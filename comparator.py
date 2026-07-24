import re
import threading
import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
import os

# ── File Reader ─────────────────────────────────

def read_file(path):
    try:
        if path.lower().endswith('.csv'):
            try:
                return pd.read_csv(path, encoding='utf-8-sig')
            except:
                return pd.read_csv(path, encoding='latin1')
        else:
            return pd.read_excel(path, engine='openpyxl')
    except Exception as e:
        raise Exception(f"Error reading file: {path}\n{e}")

# ── Normalization ─────────────────────────────────

def normalize_phone(p):
    if pd.isna(p):
        return None
    digits = re.sub(r'[^\d]', '', str(p))
    if len(digits) == 11 and digits.startswith('1'):
        digits = digits[1:]
    return digits if len(digits) >= 7 else None

def normalize_email(e):
    return str(e).strip().lower() if pd.notna(e) else ""

def normalize_name(first, last):
    return f"{str(first).strip()} {str(last).strip()}".strip()

# ── Deduplication ─────────────────────────────────

def deduplicate(df, use_phone, use_email, use_name):
    subset = []

    if use_phone and '_phone' in df:
        subset.append('_phone')
    if use_email and '_email' in df:
        subset.append('_email')
    if use_name and '_name' in df:
        subset.append('_name')

    if subset:
        df = df.drop_duplicates(subset=subset, keep='first')

    return df

# ── Match Key Builder ─────────────────────────────

def build_match_keys(df, use_phone, use_email, use_name):
    keys = []

    if use_phone:
        keys.append(df['_phone'].fillna(''))

    if use_email:
        keys.append(df['_email'].fillna(''))

    if use_name:
        keys.append(df['_name'].fillna(''))

    if not keys:
        raise Exception("Select at least one matching option")

    return pd.Series(list(zip(*keys)))

# ── Core Logic ────────────────────────────────────

def run_extraction(contact_paths, sf_paths, output_path,
                   use_phone, use_email, use_name,
                   filter_campaign, filter_adgroup,
                   mode, direction,
                   log, done):

    try:


        # ── Validation ─────────────────────

        if not output_path:
            raise Exception("Please select output file")

        if not contact_paths:
            raise Exception("Please load contact files")

        if not sf_paths:
            raise Exception("Please load Salesforce files")
        
        log("Loading Contact files...")
        contact = pd.concat([read_file(p) for p in contact_paths], ignore_index=True)
        contact['Source'] = "Contact"

        log("Loading Salesforce files...")
        sf = pd.concat([read_file(p) for p in sf_paths], ignore_index=True)
        sf['Source'] = "Salesforce"

        # Normalize
        if use_phone:
            contact['_phone'] = contact['Phone'].apply(normalize_phone)
            sf['_phone'] = sf['Phone'].apply(normalize_phone)

        if use_email:
            contact['_email'] = contact['Email'].apply(normalize_email)
            sf['_email'] = sf['Email'].apply(normalize_email)

        if use_name:
            contact['_name'] = contact.apply(lambda x: normalize_name(x.get('First Name',''), x.get('Last Name','')), axis=1)
            sf['_name'] = sf.apply(lambda x: normalize_name(x.get('First Name',''), x.get('Last Name','')), axis=1)

        # Deduplication
        contact = deduplicate(contact, use_phone, use_email, use_name)
        sf = deduplicate(sf, use_phone, use_email, use_name)

        log(f"After dedup → Contact: {len(contact)}, SF: {len(sf)}")

        # Build match keys
        contact['key'] = build_match_keys(contact, use_phone, use_email, use_name)
        sf['key'] = build_match_keys(sf, use_phone, use_email, use_name)

        contact_keys = set(contact['key'])
        sf_keys = set(sf['key'])

        # ── Comparison Modes ─────────────────────

        if mode == "UNIDIRECTIONAL":

            if direction == "CONTACT_TO_SF":
                contact['Match'] = contact['key'].isin(sf_keys)
                matched = contact[contact['Match']]
                excluded = contact[~contact['Match']]

            else:
                sf['Match'] = sf['key'].isin(contact_keys)
                matched = sf[sf['Match']]
                excluded = sf[~sf['Match']]

        else:  # BIDIRECTIONAL

            contact['Match'] = contact['key'].isin(sf_keys)
            sf['Match'] = sf['key'].isin(contact_keys)

            contact_missing = contact[~contact['Match']].copy()
            contact_missing['Missing From'] = "Salesforce"

            sf_missing = sf[~sf['Match']].copy()
            sf_missing['Missing From'] = "Contact"

            matched = pd.concat([
                contact[contact['Match']],
                sf[sf['Match']]
            ])

            excluded = pd.concat([contact_missing, sf_missing])


            # ── Optional Marketing Filters ─────────────────

        for col in ['Campaign ID', 'AdGroup ID']:

            if col not in contact:
                contact[col] = ""

            if col not in sf:
                sf[col] = ""

            contact[col] = contact[col].fillna('').astype(str).str.strip()
            sf[col] = sf[col].fillna('').astype(str).str.strip()

        # Campaign Filtering
        if filter_campaign:

            contact_campaigns = set(
                contact['Campaign ID'][contact['Campaign ID'] != '']
            )

            sf_campaigns = set(
                sf['Campaign ID'][sf['Campaign ID'] != '']
            )

            if direction == "SF_TO_CONTACT" or mode == "BIDIRECTIONAL":

                sf_missing_campaign = sf[
                    (~sf['Campaign ID'].isin(contact_campaigns)) &
                    (sf['Campaign ID'] != '')
                ].copy()

                sf_missing_campaign['Exclusion Reason'] = "Campaign ID Missing in Contact"

                excluded = pd.concat([
                    excluded,
                    sf_missing_campaign
                ], ignore_index=True)

            if direction == "CONTACT_TO_SF" or mode == "BIDIRECTIONAL":

                contact_missing_campaign = contact[
                    (~contact['Campaign ID'].isin(sf_campaigns)) &
                    (contact['Campaign ID'] != '')
                ].copy()

                contact_missing_campaign['Exclusion Reason'] = "Campaign ID Missing in SF"

                excluded = pd.concat([
                    excluded,
                    contact_missing_campaign
                ], ignore_index=True)

        # AdGroup Filtering
        if filter_adgroup:

            contact_adgroups = set(
                contact['AdGroup ID'][contact['AdGroup ID'] != '']
            )

            sf_adgroups = set(
                sf['AdGroup ID'][sf['AdGroup ID'] != '']
            )

            if direction == "SF_TO_CONTACT" or mode == "BIDIRECTIONAL":

                sf_missing_adgroup = sf[
                    (~sf['AdGroup ID'].isin(contact_adgroups)) &
                    (sf['AdGroup ID'] != '')
                ].copy()

                sf_missing_adgroup['Exclusion Reason'] = "AdGroup ID Missing in Contact"

                excluded = pd.concat([
                    excluded,
                    sf_missing_adgroup
                ], ignore_index=True)

            if direction == "CONTACT_TO_SF" or mode == "BIDIRECTIONAL":

                contact_missing_adgroup = contact[
                    (~contact['AdGroup ID'].isin(sf_adgroups)) &
                    (contact['AdGroup ID'] != '')
                ].copy()

                contact_missing_adgroup['Exclusion Reason'] = "AdGroup ID Missing in SF"

                excluded = pd.concat([
                    excluded,
                    contact_missing_adgroup
                ], ignore_index=True)

        # Remove duplicates from excluded
        excluded = excluded.drop_duplicates(subset=['key'])
        # ── Marketing Analysis (SEPARATE) ─────────

        def marketing_analysis(contact, sf):

            # Ensure columns exist
            for col in ['Campaign ID', 'AdGroup ID']:

                if col not in contact:
                    contact[col] = ""

                if col not in sf:
                    sf[col] = ""

            # Clean values
            for df in [contact, sf]:

                df['Campaign ID'] = df['Campaign ID'].fillna('').astype(str).str.strip()
                df['AdGroup ID'] = df['AdGroup ID'].fillna('').astype(str).str.strip()

            # Build comparison keys using email first
            contact_cmp = contact[['key', 'Campaign ID', 'AdGroup ID']].copy()
            sf_cmp = sf[['key', 'Campaign ID', 'AdGroup ID']].copy()

            contact_cmp.columns = ['Key', 'Contact Campaign', 'Contact AdGroup']
            sf_cmp.columns = ['Key', 'SF Campaign', 'SF AdGroup']

            # Remove empty keys
            contact_cmp = contact_cmp[
                contact_cmp['Key'].astype(str).str.strip() != ''
            ]

            sf_cmp = sf_cmp[
                sf_cmp['Key'].astype(str).str.strip() != ''
            ]

            contact_cmp = contact_cmp.drop_duplicates(subset=['Key'])
            sf_cmp = sf_cmp.drop_duplicates(subset=['Key'])

            # Merge
            merged = pd.merge(
                contact_cmp,
                sf_cmp,
                on='Key',
                how='outer'
            )

            # Campaign Status
            def campaign_status(row):

                c = row['Contact Campaign']
                s = row['SF Campaign']

                if c and s:
                    if c == s:
                        return "Match"
                    return "Different"

                elif c and not s:
                    return "Missing in SF"

                elif s and not c:
                    return "Missing in Contact"

                return "Missing Both"

            merged['Campaign Status'] = merged.apply(campaign_status, axis=1)

            # AdGroup Status
            def adgroup_status(row):

                c = row['Contact AdGroup']
                s = row['SF AdGroup']

                if c and s:
                    if c == s:
                        return "Match"
                    return "Different"

                elif c and not s:
                    return "Missing in SF"

                elif s and not c:
                    return "Missing in Contact"

                return "Missing Both"

            merged['AdGroup Status'] = merged.apply(adgroup_status, axis=1)


            # ── Unique Entity Counts ─────────────────────

            contact_unique_keys = set(contact['key'])

            sf_unique_keys = set(sf['key'])

            total = len(
                contact_unique_keys.union(sf_unique_keys)
            )

            # Detailed Summary
            marketing_summary = pd.DataFrame({

                "Metric": [

                    "Comparison Mode",
                    "Direction",

                    "Contact Records",
                    "SF Records",

                    "Total Unique Entities",

                    "Matched Unique Entities",
                    "Unmatched Unique Entities",

                    "Matched Rows",
                    "Excluded Rows",

                    "Contact Missing Records",
                    "SF Missing Records",

                    "Match Percentage",
                    "Missing Percentage"

                ],

                "Value": [

                    mode,

                    direction if mode == "UNIDIRECTIONAL"
                    else "BOTH DIRECTIONS",

                    len(contact),

                    len(sf),

                    total,

                    len(
                        set(contact['key']).intersection(
                            set(sf['key'])
                        )
                    ),

                    len(
                        set(contact['key']).symmetric_difference(
                            set(sf['key'])
                        )
                    ),

                    len(matched),

                    len(excluded),

                    (
                        len(
                            excluded[
                                excluded.get(
                                    'Missing From',
                                    ''
                                ) == 'Salesforce'
                            ]
                        )
                        if 'Missing From' in excluded.columns
                        else 0
                    ),

                    (
                        len(
                            excluded[
                                excluded.get(
                                    'Missing From',
                                    ''
                                ) == 'Contact'
                            ]
                        )
                        if 'Missing From' in excluded.columns
                        else 0
                    ),

                    round(
                        (
                            len(
                                set(contact['key']).intersection(
                                    set(sf['key'])
                                )
                            ) / total
                        ) * 100,
                        2
                    ) if total else 0,

                    round(
                        (
                            len(
                                set(contact['key']).symmetric_difference(
                                    set(sf['key'])
                                )
                            ) / total
                        ) * 100,
                        2
                    ) if total else 0

                ]

            })

            return merged, marketing_summary



        campaign_detail_df, marketing_df = marketing_analysis(contact, sf)

        # ── Summary ─────────────────────────────
        contact_rows = len(contact)
        sf_rows = len(sf)

        total_physical_rows = contact_rows + sf_rows

        unique_total_keys = len(
            set(contact['key']).union(set(sf['key']))
        )

        unique_matched_keys = len(
            set(contact['key']).intersection(set(sf['key']))
        )

        unique_excluded_keys = unique_total_keys - unique_matched_keys

        contact_matched_rows = contact['Match'].sum() if 'Match' in contact else 0
        sf_matched_rows = sf['Match'].sum() if 'Match' in sf else 0

        total_matched_rows = contact_matched_rows + sf_matched_rows

        # matched_count = unique_matched_keys
        # excluded_count = unique_excluded_keys

        summary = pd.DataFrame({

            "Metric": [

                "Mode",
                "Direction",

                "Contact Rows",
                "SF Rows",
                "Total Physical Rows",

                "Unique Combined Keys",
                "Unique Matched Keys",
                "Unique Excluded Keys",

                "Contact Matched Rows",
                "SF Matched Rows",
                "Total Matched Rows",

                "Unique Match %",
                "Unique Missing %"

            ],

            "Value": [

                mode,
                direction if mode == "UNIDIRECTIONAL" else "Two-Way Comparison",

                contact_rows,
                sf_rows,
                total_physical_rows,

                unique_total_keys,
                unique_matched_keys,
                unique_excluded_keys,

                contact_matched_rows,
                sf_matched_rows,
                total_matched_rows,

                round((unique_matched_keys / unique_total_keys) * 100, 2)
                if unique_total_keys else 0,

                round((unique_excluded_keys / unique_total_keys) * 100, 2)
                if unique_total_keys else 0

            ]

        })

        # ── Save Excel ─────────────────────────

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            contact.to_excel(writer, sheet_name='Contact Data', index=False)
            sf.to_excel(writer, sheet_name='SF Data', index=False)
            matched.to_excel(writer, sheet_name='Matched', index=False)
            excluded.to_excel(writer, sheet_name='Excluded', index=False)
            summary.to_excel(writer, sheet_name='Summary', index=False)
            campaign_detail_df.to_excel(writer, sheet_name='Campaign Comparison', index=False)
            marketing_df.to_excel(writer, sheet_name='Marketing Summary', index=False)

        log("══════════════════════════════")
        log(f"MODE: {mode}")

        matching_used = []

        if use_phone:
            matching_used.append("Phone")

        if use_email:
            matching_used.append("Email")

        if use_name:
            matching_used.append("Name")

        log(f"Matching Using: {', '.join(matching_used)}")

        log(f"Contact Records: {len(contact)}")
        log(f"SF Records: {len(sf)}")

        log(f"Unique Matched Entities: {unique_matched_keys}")
        log(f"Unique Excluded Entities: {unique_excluded_keys}")

        log(f"Matched Rows Sheet Count: {len(matched)}")
        log(f"Excluded Rows Sheet Count: {len(excluded)}")

        # Marketing summary
        try:

            campaign_match = (campaign_detail_df['Campaign Status'] == 'Match').sum()

            campaign_missing_sf = (
                campaign_detail_df['Campaign Status'] == 'Missing in SF'
            ).sum()

            campaign_missing_contact = (
                campaign_detail_df['Campaign Status'] == 'Missing in Contact'
            ).sum()

            log(f"Campaign Matches: {campaign_match}")
            log(f"Campaign Missing in SF: {campaign_missing_sf}")
            log(f"Campaign Missing in Contact: {campaign_missing_contact}")

        except:
            pass

        log("Output Saved Successfully")
        log("══════════════════════════════")

        done(True)

    except Exception as e:
        log(f"❌ {e}")
        done(False)

# ── GUI ───────────────────────────────────────────

class App:
    def __init__(self, root):
        root.title("Data Comparison Tool Pro")
        root.geometry("650x650")

        self.contact_paths = []
        self.sf_paths = []

        self.use_phone = tk.BooleanVar(value=True)
        self.use_email = tk.BooleanVar(value=True)
        self.use_name = tk.BooleanVar(value=False)

        self.filter_campaign = tk.BooleanVar(value=False)
        self.filter_adgroup = tk.BooleanVar(value=False)

        self.mode = tk.StringVar(value="UNIDIRECTIONAL")
        self.direction = tk.StringVar(value="CONTACT_TO_SF")

        frame = ttk.Frame(root, padding=10)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="📂 Input Files", font=("Arial", 12, "bold")).pack(anchor="w")

        ttk.Button(frame, text="Load Contact Files", command=self.load_contacts).pack(fill="x")
        ttk.Button(frame, text="Load Salesforce Files", command=self.load_sf).pack(fill="x")

        ttk.Label(frame, text="⚙️ Matching Options", font=("Arial", 12, "bold")).pack(anchor="w")

        ttk.Checkbutton(frame, text="Phone", variable=self.use_phone).pack(anchor="w")
        ttk.Checkbutton(frame, text="Email", variable=self.use_email).pack(anchor="w")
        ttk.Checkbutton(frame, text="Name", variable=self.use_name).pack(anchor="w")

        ttk.Label(frame, text="📢 Marketing Filters", font=("Arial", 12, "bold")).pack(anchor="w")

        ttk.Checkbutton(
            frame,
            text="Exclude Missing Campaign IDs",
            variable=self.filter_campaign
        ).pack(anchor="w")

        ttk.Checkbutton(
            frame,
            text="Exclude Missing AdGroup IDs",
            variable=self.filter_adgroup
        ).pack(anchor="w")

        ttk.Label(frame, text="🔄 Comparison Mode", font=("Arial", 12, "bold")).pack(anchor="w")

        ttk.Radiobutton(frame, text="Unidirectional", variable=self.mode, value="UNIDIRECTIONAL", command=self.toggle_direction).pack(anchor="w")
        ttk.Radiobutton(frame, text="Bidirectional", variable=self.mode, value="BIDIRECTIONAL", command=self.toggle_direction).pack(anchor="w")

        self.direction_label = ttk.Label(frame, text="➡️ Direction", font=("Arial", 12, "bold"))
        self.direction_label.pack(anchor="w")

        self.dir_contact_sf = ttk.Radiobutton(frame, text="Contact → SF", variable=self.direction, value="CONTACT_TO_SF")
        self.dir_sf_contact = ttk.Radiobutton(frame, text="SF → Contact", variable=self.direction, value="SF_TO_CONTACT")

        self.dir_contact_sf.pack(anchor="w")
        self.dir_sf_contact.pack(anchor="w")

        ttk.Button(frame, text="🗑 Clear Selected Files", command=self.clear_files).pack(fill="x")
        ttk.Button(frame, text="💾 Select Output File", command=self.save_output).pack(fill="x")
        ttk.Button(frame, text="🚀 Run Comparison", command=self.run).pack(fill="x")

        self.log_box = tk.Text(frame, height=12)
        self.log_box.pack(fill="both", expand=True)

        self.toggle_direction()

         
    def log(self, msg):
        self.log_box.insert(tk.END, msg + "\n")
        self.log_box.see(tk.END)

    def load_contacts(self):
        self.contact_paths = filedialog.askopenfilenames()
        self.log(f"{len(self.contact_paths)} Contact files loaded")

        for f in self.contact_paths:
            self.log(f"   • {os.path.basename(f)}")

    def load_sf(self):
        self.sf_paths = filedialog.askopenfilenames()
        self.log(f"{len(self.sf_paths)} Salesforce files loaded")

        for f in self.sf_paths:
            self.log(f"   • {os.path.basename(f)}")

    def save_output(self):
        self.output_path = filedialog.asksaveasfilename(defaultextension=".xlsx")

    def clear_files(self):

        self.contact_paths = []
        self.sf_paths = []

        self.output_path = ""

        self.log_box.delete("1.0", tk.END)

        self.log("✅ Cleared all selected files")

    # def run(self):
    #     threading.Thread(target=run_extraction, args=(
    #         self.contact_paths,
    #         self.sf_paths,
    #         self.output_path,
    #         self.use_phone.get(),
    #         self.use_email.get(),
    #         self.use_name.get(),
    #         self.mode.get(),
    #         self.direction.get(),
    #         self.log,
    #         lambda s: self.log("Finished" if s else "Failed")
    #     )).start()


    def run(self):

    # ── GUI Execution Summary ─────────────────

        self.log("══════════════════════════════")
        self.log(f"Mode Selected: {self.mode.get()}")

        if self.mode.get() == "UNIDIRECTIONAL":
            self.log(f"Direction: {self.direction.get()}")

        selected = []

        if self.use_phone.get():
            selected.append("Phone")

        if self.use_email.get():
            selected.append("Email")

        if self.use_name.get():
            selected.append("Name")

        self.log(f"Matching On: {', '.join(selected)}")
        self.log("══════════════════════════════")

        # ── Run Thread ───────────────────────────

        threading.Thread(target=run_extraction, args=(
            self.contact_paths,
            self.sf_paths,
            self.output_path,
            self.use_phone.get(),
            self.use_email.get(),
            self.use_name.get(),
            self.filter_campaign.get(),
            self.filter_adgroup.get(),
            self.mode.get(),
            self.direction.get(),
            self.log,
            lambda s: self.log("Finished" if s else "Failed")
        )).start()


    def toggle_direction(self):
        if self.mode.get() == "BIDIRECTIONAL":
            # Disable direction options
            self.dir_contact_sf.config(state="disabled")
            self.dir_sf_contact.config(state="disabled")

            # Grey out label
            self.direction_label.config(foreground="grey")

        else:
            # Enable direction options
            self.dir_contact_sf.config(state="normal")
            self.dir_sf_contact.config(state="normal")

            # Restore label color
            self.direction_label.config(foreground="black")

root = tk.Tk()
App(root)
root.mainloop()