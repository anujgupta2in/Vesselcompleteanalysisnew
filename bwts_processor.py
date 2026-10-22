import pandas as pd
import numpy as np
import re

# Machinery Location names vessels use for the ballast water treatment system
BWTS_LOCATION_PATTERN = r'BWTS|BWMS|Ballast Water Treatment|Ballast Treatment|Ballast Water Management'


class BWTSProcessor:
    def __init__(self):
        """Initialize BWTSProcessor."""
        self.job_code_columns = ['UI Job Code', 'Job Code', 'JobCode', 'Code']
        self.title_columns = ['Title', 'J3 Job Title', 'Task Description', 'Job Title']

    def get_bwts_data(self, data):
        """All BWTS rows."""
        return data[data['Machinery Location'].str.contains(BWTS_LOCATION_PATTERN, case=False, na=False, regex=True)].copy()

    def create_task_count_table(self, data):
        """Create task count analysis table for BWTS."""
        try:
            bwts_data = self.get_bwts_data(data)
            if bwts_data.empty:
                return pd.DataFrame()

            task_counts = bwts_data.groupby('Machinery Location')['Job Code'].count().reset_index()
            task_counts.columns = ['Machinery Location', 'Task Count']
            return task_counts.sort_values(by='Task Count', ascending=False)

        except Exception as e:
            print(f"Error creating BWTS task count table: {str(e)}")
            return pd.DataFrame()

    def find_reference_sheet(self, ref_sheet, preferred_sheet=None):
        """Name of the BWTS sheet in the reference workbook, or None.

        Order: the sheet for the selected BWTS model, then 'BWTS', then any sheet named after BWTS/ballast.
        """
        sheet_names = pd.ExcelFile(ref_sheet).sheet_names
        if preferred_sheet in sheet_names:
            return preferred_sheet
        if 'BWTS' in sheet_names:
            return 'BWTS'
        return next((s for s in sheet_names if 'bwts' in s.lower() or 'ballast' in s.lower()), None)

    def _load_reference(self, ref_sheet, preferred_sheet=None):
        """Return (reference DataFrame, job code column) or raise ValueError with a readable message."""
        sheet = self.find_reference_sheet(ref_sheet, preferred_sheet)
        if sheet is None:
            raise ValueError("No BWTS sheet found in the reference workbook")
        reference = pd.read_excel(ref_sheet, sheet_name=sheet)
        job_code_col = next((c for c in self.job_code_columns if c in reference.columns), None)
        if job_code_col is None:
            raise ValueError(f"No job code column found in reference sheet '{sheet}'. "
                             f"Available columns: {', '.join(map(str, reference.columns))}")
        reference[job_code_col] = reference[job_code_col].astype(str)
        return reference, job_code_col

    def create_reference_pivot(self, data, ref_sheet, preferred_sheet=None):
        """Reference jobs (rows) against each BWTS machinery (columns) with job counts."""
        bwts_jobs = self.get_bwts_data(data)
        bwts_jobs['Job Codecopy'] = bwts_jobs['Job Code'].astype(str)
        reference, job_code_col = self._load_reference(ref_sheet, preferred_sheet)

        merged = bwts_jobs.merge(reference, left_on='Job Codecopy', right_on=job_code_col,
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

    def process_reference_data(self, data, ref_sheet, preferred_sheet=None):
        """Reference jobs not found on the vessel's BWTS.

        Args:
            data: DataFrame containing the machinery data
            ref_sheet: Path to the reference Excel file
            preferred_sheet: Optional specific sheet name to use for BWTS model
        """
        try:
            bwts_jobs = self.get_bwts_data(data)
            reference, job_code_col = self._load_reference(ref_sheet, preferred_sheet)

            # A vessel with no BWTS jobs at all is missing every reference job
            vessel_codes = bwts_jobs['Job Code'].astype(str)
            missing = reference[~reference[job_code_col].isin(vessel_codes)]
            if 'Remarks' in missing.columns:
                missing = missing.drop(columns=['Remarks'])
            return missing.reset_index(drop=True)

        except Exception as e:
            print(f"Error processing BWTS reference data: {str(e)}")
            return pd.DataFrame({
                'Job Code': ['Error'],
                'Title': [f'Unable to process reference data: {str(e)}'],
                'Frequency': ['N/A']
            })
