import pandas as pd
import numpy as np
import re

# A purifier itself, e.g. "FO Purifier#1", "ME LO Purifier#1", "AE LO PurifierNo2"
# (not its feed pump or heater, e.g. "ME LO Purifier Feed Pump#1", "FO Purifier Heater#1")
PURIFIER_UNIT_PATTERN = r'Purifier\s*(?:#|No\.?\s*)\d+\s*$'


class PurifierProcessor:
    def __init__(self):
        """Initialize PurifierProcessor."""
        self.job_code_columns = ['UI Job Code', 'Job Code', 'JobCode', 'Code']
        self.title_columns = ['Title', 'J3 Job Title', 'Task Description', 'Job Title']

    def get_purifier_data(self, data):
        """All purifier related rows: purifiers plus their feed pumps and heaters."""
        return data[data['Machinery Location'].str.contains('Purifier', case=False, na=False)].copy()

    def extract_running_hours(self, data):
        """Extract running hours for each purifier."""
        try:
            if 'Machinery Location' not in data.columns or 'Machinery Running Hours' not in data.columns:
                return pd.DataFrame(columns=['Purifier', 'Running Hours'])

            purifier_data = data[data['Machinery Location'].str.contains(
                PURIFIER_UNIT_PATTERN, case=False, na=False, regex=True)].copy()
            if purifier_data.empty:
                return pd.DataFrame(columns=['Purifier', 'Running Hours'])

            purifier_data['Hours'] = pd.to_numeric(purifier_data['Machinery Running Hours'], errors='coerce')

            def most_common(hours):
                hours = hours.dropna()
                return hours.mode().max() if not hours.empty else np.nan

            # Most common value per purifier, since a few rows can carry other counters
            running_hours = (purifier_data.groupby('Machinery Location')['Hours']
                             .agg(most_common).reset_index()
                             .rename(columns={'Machinery Location': 'Purifier', 'Hours': 'Running Hours'}))
            running_hours['Running Hours'] = running_hours['Running Hours'].apply(
                lambda x: f"{int(x):,}" if pd.notnull(x) and x != 0 else "N/A"
            )
            return running_hours

        except Exception as e:
            print(f"Error extracting purifier running hours: {str(e)}")
            return pd.DataFrame({
                'Purifier': ['Error'],
                'Running Hours': [f'Error: {str(e)}']
            })

    def create_task_count_table(self, data):
        """Create task count analysis table for purifiers."""
        try:
            purifier_data = self.get_purifier_data(data)
            if purifier_data.empty:
                return pd.DataFrame()

            task_counts = purifier_data.groupby('Machinery Location').size().reset_index(name='Task Count')
            task_counts['Machinery Location'] = task_counts['Machinery Location'].apply(
                lambda x: str(x).replace('  ', ' ').strip()
            )
            return task_counts.sort_values('Machinery Location').reset_index(drop=True)

        except Exception as e:
            print(f"Error creating purifier task count table: {str(e)}")
            return pd.DataFrame()

    def find_reference_sheet(self, ref_sheet):
        """Name of the purifier sheet in the reference workbook, or None."""
        sheet_names = pd.ExcelFile(ref_sheet).sheet_names
        if 'Purifiers' in sheet_names:
            return 'Purifiers'
        # Case-insensitive fallback; never match on "PU" alone, which also matches "Pumps"
        return next((s for s in sheet_names if 'purifier' in s.lower()), None)

    def _load_reference(self, ref_sheet):
        """Return (reference DataFrame, job code column) or raise ValueError with a readable message."""
        sheet = self.find_reference_sheet(ref_sheet)
        if sheet is None:
            raise ValueError("No 'Purifiers' sheet found in the reference workbook")
        reference = pd.read_excel(ref_sheet, sheet_name=sheet)
        job_code_col = next((c for c in self.job_code_columns if c in reference.columns), None)
        if job_code_col is None:
            raise ValueError(f"No job code column found in reference sheet '{sheet}'. "
                             f"Available columns: {', '.join(map(str, reference.columns))}")
        reference[job_code_col] = reference[job_code_col].astype(str)
        return reference, job_code_col

    def create_reference_pivot(self, data, ref_sheet):
        """Reference jobs (rows) against each purifier machinery (columns) with job counts."""
        purifier_jobs = self.get_purifier_data(data)
        purifier_jobs['Job Codecopy'] = purifier_jobs['Job Code'].astype(str)
        reference, job_code_col = self._load_reference(ref_sheet)

        merged = purifier_jobs.merge(reference, left_on='Job Codecopy', right_on=job_code_col,
                                     suffixes=('_filtered', '_ref'))
        title_col = next((c for c in self.title_columns if c in merged.columns), None)
        if merged.empty or title_col is None:
            return pd.DataFrame()

        return merged.pivot_table(
            index=title_col,
            columns='Machinery Location',
            values='Job Codecopy',
            aggfunc='count'
        ).fillna(0).astype(int)

    def process_reference_data(self, data, ref_sheet):
        """Reference jobs not found on any purifier machinery."""
        try:
            purifier_jobs = self.get_purifier_data(data)
            reference, job_code_col = self._load_reference(ref_sheet)
            if purifier_jobs.empty or reference.empty:
                return pd.DataFrame({
                    'Job Code': ['No data found'],
                    'Title': ['No matching data between current and reference'],
                    'Frequency': ['N/A']
                })

            vessel_codes = purifier_jobs['Job Code'].astype(str)
            missing = reference[~reference[job_code_col].isin(vessel_codes)]
            if 'Remarks' in missing.columns:
                missing = missing.drop(columns=['Remarks'])
            return missing.reset_index(drop=True)

        except Exception as e:
            print(f"Error processing purifier reference data: {str(e)}")
            return pd.DataFrame({
                'Job Code': ['Error'],
                'Title': [f'Unable to process reference data: {str(e)}'],
                'Frequency': ['N/A']
            })
