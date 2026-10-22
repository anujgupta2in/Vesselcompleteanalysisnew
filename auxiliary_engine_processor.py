import re

import pandas as pd
import numpy as np

# Matches the ways vessels name an auxiliary engine, e.g. "Auxiliary Engine#1",
# "Auxiliary EngineNo1", "Auxiliary Engine No.1", "Aux. Engine 1", "AE#1", "A/E No.1"
AE_NAME_PATTERN = re.compile(
    r'^\s*(?:aux(?:iliary|\.)?\s*engine|a\s*/?\s*e)\s*[-#:]?\s*(?:no\.?)?\s*[-#:]?\s*(\d+)\s*$',
    re.IGNORECASE
)


class AuxiliaryEngineProcessor:
    def __init__(self):
        """Initialize AuxiliaryEngineProcessor with component list."""
        self.components_to_check = [
            'Turbocharger - AE',
            'Air Cooler - AE',
            'Alternator - AE',
            'Connecting Rod - AE',
            'Cylinder Head - AE',
            'Fuel Injection Pump - AE',
            'Fuel Valve - AE',
            'Governor - AE',
            'Main Bearing - AE',
            'Piston - AE',
            'Attached LO Pump - AE',
            'FO Filter - AE',
            'LO Cooler - AE',
            'LO Filter - AE',
            'LT Attached CW Pump - AE',
        ]
        self.job_codes = [
            ([6619, 2157], "Piston Overhaul", r'Piston - AE#(\d+)'),
            ([2222, 2223, 2224], "Cylinder Head Overhaul", r'Cylinder Head - AE#(\d+)'),
            ([2196, 2191], "Fuel Valve Overhaul", r'Fuel Valve - AE#(\d+)'),
            (2202, "Fuel Pump Overhaul", r'Fuel Injection Pump - AE#(\d+)'),
            ([3878, 2127, 2126], "TurboCharger Overhaul", r'Auxiliary Engine#(\d+)')
        ]
        # Problems found while processing, shown as warnings in the app
        self.errors = []

    def _log_error(self, message):
        print(message)
        if message not in self.errors:
            self.errors.append(message)

    def get_ae_data(self, data):
        """Return only auxiliary engine rows with standardised naming.

        Machinery Location becomes "Auxiliary Engine#N" and Sub Component Location keeps only
        the part after the last ">" so vessels with different naming styles line up.
        """
        if 'Machinery Location' not in data.columns:
            return pd.DataFrame(columns=list(data.columns) + ['AE No'])

        engine_no = data['Machinery Location'].astype(str).str.extract(AE_NAME_PATTERN, expand=False)
        ae_df = data[engine_no.notna()].copy()
        ae_df['AE No'] = engine_no[engine_no.notna()].astype(int)
        ae_df['Machinery Location'] = 'Auxiliary Engine#' + ae_df['AE No'].astype(str)
        if 'Sub Component Location' in ae_df.columns:
            sub = ae_df['Sub Component Location']
            ae_df['Sub Component Location'] = sub.where(
                sub.isna(), sub.astype(str).str.split('>').str[-1].str.strip()
            )
        return ae_df

    def get_engine_numbers(self, ae_df):
        """Auxiliary engine numbers found in the data (e.g. [1, 2, 3, 4])."""
        if ae_df.empty:
            return []
        return sorted(ae_df['AE No'].unique().tolist())

    def extract_running_hours(self, data):
        """Extract running hours for every auxiliary engine found in the data."""
        # AE1-AE3 are always present so existing reports keep working
        aux_running_hours = {'AE1': "Not Available", 'AE2': "Not Available", 'AE3': "Not Available"}
        ae_df = self.get_ae_data(data)
        for ae_num in self.get_engine_numbers(ae_df):
            aux_running_hours.setdefault(f'AE{ae_num}', "Not Available")
            if 'Machinery Running Hours' not in ae_df.columns:
                continue
            running_hours = pd.to_numeric(
                ae_df.loc[ae_df['AE No'] == ae_num, 'Machinery Running Hours'], errors='coerce'
            ).dropna()
            if not running_hours.empty:
                aux_running_hours[f'AE{ae_num}'] = str(int(running_hours.max()))
        return aux_running_hours

    def format_unit_data(self, unit_data):
        """Format unit data for display, showing the most critical unit (cylinder)."""
        try:
            remaining = pd.to_numeric(unit_data['Remaining Running Hours'], errors='coerce')
            # PMS exports dates as dd-mm-yyyy
            last_done = pd.to_datetime(unit_data['Last Done Date'], errors='coerce', dayfirst=True)
            # Most critical = lowest remaining hours, then oldest last done date
            ordered = unit_data.assign(_remaining=remaining, _last_done=last_done).sort_values(
                ['_remaining', '_last_done'], na_position='last'
            )
            worst = ordered.iloc[0]

            def show(value, empty):
                if pd.isna(value) or str(value).strip() == '':
                    return empty
                if isinstance(value, float) and value.is_integer():
                    return str(int(value))
                return str(value)

            text = (f"Date: {show(worst['Last Done Date'], 'No Date')}\n"
                    f"RH: {show(worst['Last Done Running Hours'], 'No RH')}\n"
                    f"Remaining Hours: {show(worst['Remaining Running Hours'], 'No RH')}")

            # Per-cylinder jobs: say how many units and which one is the most critical
            units = unit_data['Sub Component Location'].dropna().nunique()
            if units > 1:
                if ordered['_remaining'].nunique(dropna=False) == 1:
                    text += f"\nUnits: {units} (all same)"
                else:
                    text += f"\nUnits: {units} (most critical: {show(worst['Sub Component Location'], '-')})"
            return text
        except Exception:
            return "Data formatting error"

    def process_job_code(self, data, job_codes, job_description, unit_pattern, engine_numbers=None):
        """Process job codes for auxiliary engines."""
        if not isinstance(job_codes, list):
            job_codes = [job_codes]
        try:
            ae_df = self.get_ae_data(data)
            engine_numbers = engine_numbers or self.get_engine_numbers(ae_df) or [1, 2, 3]
            job_data = ae_df[ae_df['Job Code'].isin(job_codes)]

            structured_data = pd.DataFrame({
                'Job Title': [job_description],
                'Frequency': [job_data['Frequency'].iloc[0] if not job_data.empty else "No Frequency"]
            })

            for unit in engine_numbers:
                unit_data = job_data[job_data['AE No'] == unit]
                if not unit_data.empty:
                    structured_data[f'AE{unit}'] = [self.format_unit_data(unit_data)]
                else:
                    structured_data[f'AE{unit}'] = ["No Data Available"]

            return structured_data
        except Exception as e:
            self._log_error(f"Error processing AE job code {job_codes}: {str(e)}")
            return pd.DataFrame({'Job Title': [job_description], 'Frequency': ["Error processing data"]})

    def get_maintenance_data(self, data):
        """Get auxiliary engine maintenance data."""
        engine_numbers = self.get_engine_numbers(self.get_ae_data(data)) or [1, 2, 3]
        maintenance_data = []
        for job_info in self.job_codes:
            if len(job_info) == 3:
                job_codes, job_description, unit_pattern = job_info
                structured_data = self.process_job_code(data, job_codes, job_description, unit_pattern, engine_numbers)
                maintenance_data.append(structured_data)

        if maintenance_data:
            return pd.concat(maintenance_data, ignore_index=True)
        return pd.DataFrame()

    def analyze_components(self, data):
        """Analyze component presence for each auxiliary engine."""
        ae_df = self.get_ae_data(data)
        engine_numbers = self.get_engine_numbers(ae_df)

        rows = []
        missing_count = 0
        for component in self.components_to_check:
            row = {'Component': component}
            missing_on = []
            for ae_num in engine_numbers:
                locations = ae_df.loc[ae_df['AE No'] == ae_num, 'Sub Component Location'].dropna().astype(str)
                found = locations.str.contains(component, regex=False).any()
                row[f'Auxiliary Engine#{ae_num}'] = "Present" if found else "Missing"
                if not found:
                    missing_on.append(f'AE#{ae_num}')

            if not engine_numbers or len(missing_on) == len(engine_numbers):
                row['Status'] = "Missing"
            elif missing_on:
                row['Status'] = f"Missing on {', '.join(missing_on)}"
            else:
                row['Status'] = "Present"
            if row['Status'] != "Present":
                missing_count += 1
            rows.append(row)

        return pd.DataFrame(rows), missing_count

    def create_task_count_table(self, data):
        """Create task count analysis table for auxiliary engines."""
        try:
            auxiliary_engine_df = self.get_ae_data(data)

            # Create basic task count pivot table
            pivot_table = auxiliary_engine_df.pivot_table(
                index='Machinery Location',
                values='Title',
                aggfunc='count'
            ).reset_index()

            # Rename columns for clarity
            pivot_table.columns = ['Machinery Location', 'Task Count']

            return pivot_table

        except Exception as e:
            self._log_error(f"Error creating task count table: {e}")
            return None

    def create_component_distribution(self, data):
        """Create component distribution analysis table."""
        if 'Sub Component Location' not in data.columns:
            return None

        ae_df = self.get_ae_data(data)
        pattern = '|'.join(re.escape(c) for c in self.components_to_check)
        filtered_df = ae_df[ae_df['Sub Component Location'].str.contains(pattern, na=False, case=False)]
        if filtered_df.empty:
            return None

        pivot_table = filtered_df.pivot_table(
            index=['Sub Component Location', 'Title'],
            columns='Machinery Location',
            values='Job Code',
            aggfunc='count',
            fill_value=0
        )
        # Engines with no matching jobs at all still get a column of zeros
        engine_columns = [f'Auxiliary Engine#{n}' for n in self.get_engine_numbers(ae_df)]
        pivot_table = pivot_table.reindex(columns=engine_columns, fill_value=0).reset_index()
        pivot_table.columns.name = None
        return pivot_table

    def find_job_gaps(self, data):
        """Compare sister auxiliary engines and list jobs missing or duplicated on an engine."""
        distribution = self.create_component_distribution(data)
        columns = ['Auxiliary Engine', 'Sub Component Location', 'Title', 'Issue']
        if distribution is None or distribution.empty:
            return pd.DataFrame(columns=columns)

        engine_columns = [c for c in distribution.columns if c.startswith('Auxiliary Engine#')]
        gaps = []
        for _, row in distribution.iterrows():
            present_on = [c.replace('Auxiliary Engine', 'AE') for c in engine_columns if row[c] > 0]
            for engine in engine_columns:
                if row[engine] == 0:
                    issue = f"Missing (present on {', '.join(present_on)})"
                elif row[engine] > 1:
                    issue = f"Duplicate ({int(row[engine])} jobs)"
                else:
                    continue
                gaps.append({
                    'Auxiliary Engine': engine,
                    'Sub Component Location': row['Sub Component Location'],
                    'Title': row['Title'],
                    'Issue': issue,
                })

        if not gaps:
            return pd.DataFrame(columns=columns)
        return pd.DataFrame(gaps).sort_values(['Auxiliary Engine', 'Sub Component Location', 'Title']).reset_index(drop=True)

    def process_reference_data(self, data, ref_sheet):
        """Process reference data for auxiliary engines."""
        try:
            filtered_df = self.get_ae_data(data)
            filtered_df['Job Codecopy'] = filtered_df['Job Code'].astype(str)
            engine_names = [f'Auxiliary Engine#{n}' for n in self.get_engine_numbers(filtered_df)]

            ref_df = pd.read_excel(ref_sheet, sheet_name='AE Jobs')
            ref_df['UI Job Code'] = ref_df['UI Job Code'].astype(str)

            result_df = filtered_df.merge(
                ref_df,
                left_on='Job Codecopy',
                right_on='UI Job Code',
                suffixes=('_filtered', '_ref')
            )

            pivot_table = result_df.pivot_table(
                index='Title',
                columns='Machinery Location',
                values='Job Codecopy',
                aggfunc='count',
                fill_value=0
            )
            pivot_table = pivot_table.reindex(columns=engine_names, fill_value=0).reset_index()
            pivot_table.columns.name = None

            # A reference job is missing if any engine lacks it, not only when the whole vessel lacks it
            engines_with_job = filtered_df.groupby('Job Codecopy')['Machinery Location'].agg(set)
            missing_on = ref_df['UI Job Code'].map(
                lambda code: [e for e in engine_names if e not in engines_with_job.get(code, set())]
            )
            if not engine_names:
                missing_on = ref_df['UI Job Code'].map(lambda code: ['All engines'])
            missing_jobs = ref_df[missing_on.str.len() > 0].copy()
            missing_jobs['Missing On'] = missing_on[missing_on.str.len() > 0].map(
                lambda engines: "All engines" if not engine_names or len(engines) == len(engine_names)
                else ', '.join(e.replace('Auxiliary Engine', 'AE') for e in engines)
            )
            if 'Remarks' in missing_jobs.columns:
                missing_jobs = missing_jobs.drop(columns=['Remarks'])
            missing_jobs.reset_index(drop=True, inplace=True)

            return pivot_table, missing_jobs
        except Exception as e:
            self._log_error(f"Error processing AE reference data: {e}")
            return None, None
