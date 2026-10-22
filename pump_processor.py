import pandas as pd
import numpy as np

def is_pump_machinery(machinery_location):
    """True for pumps, but not "Pump Room", "Pump Room Davit" or "Cargo Pumping - Pipes, Valves and Fittings".

    Two plain checks rather than one look-ahead regex: pandas 3 runs .str regexes through
    PyArrow (RE2), which does not support look-ahead.
    """
    # Remove "Pump Room"/"Pumping", then look for "pump": "Pump Room Bilge Pump" is still a pump
    remaining = machinery_location.astype(str).str.replace(r'pump\s*room|pumping', '', case=False, regex=True)
    return remaining.str.contains('pump', case=False, na=False)

# Jobs every motor-driven pump is expected to have (after mapping vessel variants onto these codes)
STANDARD_PUMP_JOBS = {
    '425': 'Pump - Minor Maintenance',
    '426': 'Pump - Major Maintenance',
    '602': 'Motor & Starter - Minor Maintenance',
    '603': 'Motor - Major Maintenance',
}


class PumpSystemProcessor:
    def __init__(self):
        self.filtered_dfpump = pd.DataFrame()
        self.pivot_table_pump = pd.DataFrame()
        self.pivot_table_resultpumpJobs = pd.DataFrame()
        self.pump_job_gaps = pd.DataFrame()
        self.sub_pump_job_gaps = pd.DataFrame()
        self.styled_pivot_table_pump = None
        self.styled_pivot_table_resultpumpJobs = None
        self.mappings = {'4406': '425', '428': '425', '2329': '425', '4656': '426', '7039': '426', '6001': '602'}

    def _mapped_codes(self, job_codes):
        codes = job_codes.astype(str).str.strip()
        return codes.map(self.mappings).fillna(codes)

    def _gap_table(self, units, required, unit_label):
        """One row per pump missing any required job: which ones, by code and title."""
        rows = []
        for unit, codes in units.items():
            missing = [c for c in required if c not in codes]
            if missing:
                rows.append({
                    unit_label: unit,
                    'Missing Jobs': ', '.join(f"{STANDARD_PUMP_JOBS[c]} ({c})" for c in missing),
                    'Missing Count': len(missing),
                    # e.g. cargo pumps with their own job set have none of the standard jobs
                    'Note': 'No standard pump jobs at all - check if this pump uses its own job set'
                            if len(missing) == len(required) else '',
                })
        columns = [unit_label, 'Missing Jobs', 'Missing Count', 'Note']
        return pd.DataFrame(rows, columns=columns)

    def find_standard_job_gaps(self, df, dfpump=None):
        """Pumps missing any of the standard pump/motor jobs.

        Returns (pumps that are machinery, pumps inside other machinery). The required set is the
        standard jobs that are also in the reference sheet, when the sheet lists any of them.
        """
        required = list(STANDARD_PUMP_JOBS)
        if dfpump is not None and 'UI Job Code' in dfpump.columns:
            in_reference = [c for c in required if c in set(dfpump['UI Job Code'].astype(str).str.strip())]
            required = in_reference or required

        data = df.assign(_code=self._mapped_codes(df['Job Code']))
        is_pump = is_pump_machinery(data['Machinery Location'])

        machinery_pumps = data[is_pump].groupby('Machinery Location')['_code'].agg(set)

        # Pumps that are a sub-component, e.g. "Feed Pump - AE LO Purifier#1 > Motor - Feed Pump#1"
        others = data[~is_pump].copy()
        others['_part'] = others['Sub Component Location'].astype(str).str.split('>').str[0].str.strip()
        others = others[others['_part'].str.contains(r'\bpump\b', case=False, regex=True)]
        others['_unit'] = others['Machinery Location'] + ' > ' + others['_part']
        sub_pumps = others.groupby('_unit')['_code'].agg(set)
        # Only pumps that follow the standard set; engine-driven pumps (fuel injection, attached LO/CW pumps)
        # have their own engine jobs and no motor
        sub_pumps = sub_pumps[sub_pumps.map(lambda codes: bool(codes & {'425', '426'}))]

        return (self._gap_table(machinery_pumps, required, 'Pump'),
                self._gap_table(sub_pumps, required, 'Pump (inside other machinery)'))

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def process_pump_data(self, df, dfpump):
        try:
            print("🚀 Starting pump data processing...")

            if 'UI Job Code' not in dfpump.columns:
                raise ValueError("Reference sheet missing 'UI Job Code' column")

            # 🔹 Pump Location Count Table
            pump_df = df[is_pump_machinery(df['Machinery Location'])]
            self.pivot_table_pump = pump_df.pivot_table(index='Machinery Location', values='Title', aggfunc='count')

            # ✅ Red highlight for counts > 6
            self.styled_pivot_table_pump = self.pivot_table_pump.style.apply(
                lambda x: ['background-color: red; color: white;' if v > 6 else '' for v in x], axis=1
            ).set_table_styles([
                {'selector': 'th', 'props': [('font-weight', 'bold')]},
                {'selector': 'td', 'props': [('text-align', 'center'), ('min-width', '150px')]},
                {'selector': 'td:first-child', 'props': [('text-align', 'left')]}
            ], overwrite=False)

            print("✅ Pump count table created")

            # 🔹 Pumps missing standard jobs (independent of the reference matching below)
            self.pump_job_gaps, self.sub_pump_job_gaps = self.find_standard_job_gaps(df, dfpump)

            # 🔹 Filter and map job codes
            self.filtered_dfpump = pump_df.copy()

            self.filtered_dfpump['Job Codecopy'] = self.filtered_dfpump['Job Code'].astype(str).str.strip()
            self.filtered_dfpump['Job Codecopy'] = self.filtered_dfpump['Job Codecopy'].apply(self.safe_convert_to_string)

            self.filtered_dfpump['Job Code'] = self.filtered_dfpump['Job Code'].astype(str)
            self.filtered_dfpump['Job Code'] = self.filtered_dfpump['Job Code'].map(self.mappings).fillna(self.filtered_dfpump['Job Code'])
            self.filtered_dfpump['Job Code'] = self.filtered_dfpump['Job Code'].astype(str)

            job_codespump = dfpump['UI Job Code'].astype(str).tolist()
            matching_jobspump = self.filtered_dfpump[self.filtered_dfpump['Job Code'].isin(job_codespump)]

            # 🔹 Create the result pivot table
            pivot_table_resultpumpJobs = matching_jobspump.pivot_table(
                index=['Machinery Location'],
                columns='Job Code',
                values='Title',
                aggfunc='count',
                margins=True,
                margins_name='Total'
            )

            if pivot_table_resultpumpJobs.empty:
                print("⚠️ No matching pump jobs found.")
                self.pivot_table_resultpumpJobs = pd.DataFrame()
                self.styled_pivot_table_resultpumpJobs = None
                return self.pivot_table_resultpumpJobs

            pivot_table_resultpumpJobs.fillna(0, inplace=True)
            pivot_table_resultpumpJobs = pivot_table_resultpumpJobs.astype(int)
            self.pivot_table_resultpumpJobs = pivot_table_resultpumpJobs.copy()

            print("✅ Mapped Job Code Summary pivot created")

            # ✅ Green highlight for non-zero job counts
            styled = self.pivot_table_resultpumpJobs.style.highlight_between(
                left=1, right=9999, props="background-color: #d4edda; color: black;"
            )

            styled = styled.set_table_styles([
                {'selector': 'th', 'props': [('font-weight', 'bold')]},
                {'selector': 'td', 'props': [('text-align', 'center'), ('min-width', '200px')]},
                {'selector': 'td:first-child', 'props': [('text-align', 'left')]}
            ], overwrite=False)

            self.styled_pivot_table_resultpumpJobs = styled

            print("✅ Styling applied safely with highlight.")
            return self.pivot_table_resultpumpJobs

        except Exception as e:
            import traceback
            traceback.print_exc()
            return pd.DataFrame({'Error': [f'Pump data processing failed: {str(e)}']})
