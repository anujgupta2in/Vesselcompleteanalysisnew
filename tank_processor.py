import pandas as pd
import numpy as np

class TankSystemProcessor:
    def __init__(self):
        self.filtered_dftanksjobs = pd.DataFrame()
        self.result_dftanks = pd.DataFrame()
        self.pivot_table_resulttanksJobs = pd.DataFrame()
        self.styled_pivot_table_resulttanksJobs = None
        self.missingjobstankresult = pd.DataFrame()

    def safe_convert_to_string(self, value):
        try:
            return str(value).strip()
        except Exception:
            return ''

    def format_blank(self, val):
        return "" if val == -1 else val

    def process_tank_data(self, df, dftanks):
        try:
            tanks = [
                'Ballast System','Bilge and Sludge System','Fuel Oil Service System','Cargo Handling System',
                'Lubricating Oil Purification System','Fuel Oil Storage and Transfer System','Fresh Water System',
                'Cargo Ventilation System','Lubricating Oil Storage and Transfer System','Lubricating Oil Service System',
                'Steam and Condensate System','Cooling Fresh Water System','Fuel Oil Purification System',
                'Cooling Sea Water System','Stern Tube System','Waste Handling'
            ]

            # Check if Function column exists (case-insensitive fallback)
            func_col = 'Function'
            if func_col not in df.columns:
                for col in df.columns:
                    if col.lower() == 'function':
                        func_col = col
                        break
            
            if func_col in df.columns:
                self.filtered_dftanksjobs = df[df[func_col].astype(str).str.contains('|'.join(tanks), na=False, case=False)].copy()
            else:
                self.filtered_dftanksjobs = pd.DataFrame(columns=df.columns)

            if 'Job Codecopy' not in self.filtered_dftanksjobs.columns:
                if 'Job Code' in self.filtered_dftanksjobs.columns:
                    self.filtered_dftanksjobs['Job Codecopy'] = self.filtered_dftanksjobs['Job Code'].astype(str)
                else:
                    self.filtered_dftanksjobs['Job Codecopy'] = ''
            
            self.filtered_dftanksjobs['Job Codecopy'] = self.filtered_dftanksjobs['Job Codecopy'].astype(object)
            self.filtered_dftanksjobs['Job Codecopy'] = self.filtered_dftanksjobs['Job Codecopy'].apply(self.safe_convert_to_string)

            # Detect UI Job Code column in dftanks
            ref_job_col = 'UI Job Code'
            for col in ['UI Job Code', 'Job Code', 'JobCode', 'Code']:
                if col in dftanks.columns:
                    ref_job_col = col
                    break
            
            if ref_job_col in dftanks.columns:
                dftanks[ref_job_col] = dftanks[ref_job_col].apply(self.safe_convert_to_string)
            else:
                raise ValueError("No suitable UI Job Code column found in reference sheet.")

            # Perform merge
            self.result_dftanks = self.filtered_dftanksjobs.merge(
                dftanks,
                left_on='Job Codecopy',
                right_on=ref_job_col,
                suffixes=('_filtered', '_ref')
            )
            self.result_dftanks.reset_index(drop=True, inplace=True)

            possible_titles = ['Title', 'J3 Job Title', 'Task Description', 'Job Title']
            title_col = next((col for col in possible_titles if col in self.result_dftanks.columns), None)
            
            if title_col is None and not self.result_dftanks.empty:
                title_col = self.result_dftanks.columns[0]

            if title_col is None:
                pivot_table = pd.DataFrame()
            else:
                pivot_col = 'Function' if 'Function' in self.result_dftanks.columns else func_col
                pivot_table = self.result_dftanks.pivot_table(
                    index=title_col,
                    columns=pivot_col,
                    values='Job Codecopy',
                    aggfunc='count'
                )
                pivot_table.replace(np.nan, '', inplace=True)
                pivot_table.replace('', -1, inplace=True)
                pivot_table = pivot_table.astype(int)
                
                # Standard and robust mapping that works across all Pandas versions
                pivot_table = pivot_table.apply(lambda col: col.map(self.format_blank))

            self.pivot_table_resulttanksJobs = pivot_table
            
            if not self.pivot_table_resulttanksJobs.empty:
                self.styled_pivot_table_resulttanksJobs = self.pivot_table_resulttanksJobs.style\
                    .set_table_styles([
                        {'selector': 'th', 'props': [('font-weight', 'bold'), ('text-align', 'left')]},
                        {'selector': 'td', 'props': [('text-align', 'left'), ('min-width', '120px')]},
                        {'selector': 'td:first-child', 'props': [('text-align', 'left'), ('min-width', '250px')]}
                    ], overwrite=False)\
                    .set_table_attributes("class='dataframe' style='margin-left: 0 !important; margin-right: auto; width: 100%'")
            else:
                self.styled_pivot_table_resulttanksJobs = None

            self.missingjobstankresult = dftanks[~dftanks[ref_job_col].isin(
                self.filtered_dftanksjobs['Job Codecopy']
            )].copy()
            
            if 'Remarks' in self.missingjobstankresult.columns:
                self.missingjobstankresult.drop(columns=['Remarks'], inplace=True)
            self.missingjobstankresult.reset_index(drop=True, inplace=True)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.pivot_table_resulttanksJobs = pd.DataFrame({'Error': [f'Tank data processing failed: {str(e)}']})
            self.missingjobstankresult = pd.DataFrame()
