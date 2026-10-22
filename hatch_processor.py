import pandas as pd
import numpy as np
import re

# Machinery Location names for hatch covers and their hydraulics
HATCH_LOCATION_PATTERN = r'Hatch|Cargo Opening'


def natural_key(series):
    """Sort key so "Hatch Cover#10" comes after "Hatch Cover#9", not after "Hatch Cover#1"."""
    return series.astype(str).map(
        lambda s: tuple(int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', s))
    )


class HatchProcessor:
    def __init__(self):
        """Initialize HatchProcessor."""
        self.job_code_columns = ['UI Job Code', 'Job Code', 'JobCode', 'Code']
        self.title_columns = ['J3 Job Title', 'Job Title', 'Title', 'Description']

    def get_hatch_data(self, data):
        """All hatch cover rows."""
        return data[data['Machinery Location'].str.contains(HATCH_LOCATION_PATTERN, case=False, na=False, regex=True)].copy()

    def create_task_count_table(self, data):
        """Create task count analysis table for Hatches."""
        try:
            hatch_data = self.get_hatch_data(data)
            if hatch_data.empty:
                return pd.DataFrame()

            task_counts = hatch_data.groupby('Machinery Location')['Job Code'].count().reset_index()
            task_counts.columns = ['Machinery Location', 'Task Count']
            return task_counts.sort_values('Machinery Location', key=natural_key).reset_index(drop=True)

        except Exception as e:
            print(f"Error creating Hatch task count table: {str(e)}")
            return pd.DataFrame()

    def find_reference_sheet(self, ref_sheet, preferred_sheet='Hatch'):
        """Name of the hatch sheet in the reference workbook, or None."""
        sheet_names = pd.ExcelFile(ref_sheet).sheet_names
        if preferred_sheet in sheet_names:
            return preferred_sheet
        if 'Hatch' in sheet_names:
            return 'Hatch'
        # Only sheets named after hatches; "Cargo" alone would also match the cargo handling sheet
        return next((s for s in sheet_names if 'hatch' in s.lower()), None)

    def _load_reference(self, ref_sheet, preferred_sheet='Hatch'):
        """Return (reference DataFrame, job code column) or raise ValueError with a readable message."""
        sheet = self.find_reference_sheet(ref_sheet, preferred_sheet)
        if sheet is None:
            raise ValueError("No 'Hatch' sheet found in the reference workbook")
        reference = pd.read_excel(ref_sheet, sheet_name=sheet)
        job_code_col = next((c for c in self.job_code_columns if c in reference.columns), None)
        if job_code_col is None:
            raise ValueError(f"No job code column found in reference sheet '{sheet}'. "
                             f"Available columns: {', '.join(map(str, reference.columns))}")
        reference[job_code_col] = reference[job_code_col].astype(str)
        return reference, job_code_col

    def process_reference_data(self, data, ref_sheet, preferred_sheet='Hatch'):
        """Reference hatch jobs not found on the vessel.

        Returns an empty DataFrame when the vessel has no hatch covers (e.g. a tanker),
        so it is not counted as missing jobs.
        """
        try:
            hatch_jobs = self.get_hatch_data(data)
            if hatch_jobs.empty:
                return pd.DataFrame()

            reference, job_code_col = self._load_reference(ref_sheet, preferred_sheet)
            missing = reference[~reference[job_code_col].isin(hatch_jobs['Job Code'].astype(str))]
            if 'Remarks' in missing.columns:
                missing = missing.drop(columns=['Remarks'])
            return missing.reset_index(drop=True)

        except Exception as e:
            print(f"Error processing Hatch reference data: {str(e)}")
            return pd.DataFrame({
                'Job Code': ['Error'],
                'Title': [f'Unable to process reference data: {str(e)}'],
                'Source': ['N/A']
            })

    def create_reference_pivot_table(self, data, ref_sheet, preferred_sheet='Hatch'):
        """Reference jobs (rows) against each hatch machinery (columns) with job counts."""
        try:
            hatch_jobs = self.get_hatch_data(data)
            if hatch_jobs.empty:
                return pd.DataFrame()
            hatch_jobs['Job Codecopy'] = hatch_jobs['Job Code'].astype(str)
            reference, job_code_col = self._load_reference(ref_sheet, preferred_sheet)

            # Prefer the reference sheet's title; fall back to its second column
            title_col = next((c for c in self.title_columns if c in reference.columns), None)
            if title_col is None and len(reference.columns) > 1:
                title_col = reference.columns[1]

            merged = hatch_jobs.merge(reference, left_on='Job Codecopy', right_on=job_code_col,
                                      suffixes=('_filtered', '_ref'))
            if title_col == 'Title':
                title_col = 'Title_ref'  # both sheets have "Title"; use the reference one
            if merged.empty or title_col not in merged.columns:
                return pd.DataFrame()

            pivot = merged.pivot_table(
                index=title_col,
                columns='Machinery Location',
                values='Job Codecopy',
                aggfunc='count'
            ).fillna(0).astype(int)
            return pivot[sorted(pivot.columns, key=lambda c: natural_key(pd.Series([c]))[0])]

        except Exception as e:
            print(f"Error creating Hatch reference pivot table: {str(e)}")
            return pd.DataFrame()
