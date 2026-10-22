import pandas as pd
from datetime import datetime

class CSVValidator:
    def __init__(self):
        # Define required columns and their possible alternative names
        self.column_mappings = {
            'Vessel': ['Vessel', 'Ship Name', 'Vessel Name'],
            'Job Code': ['Job Code', 'CMS Code', 'Work Code', 'Task Code'],
            'Frequency': ['Frequency', 'Interval', 'Maintenance Frequency'],
            'Calculated Due Date': ['Calculated Due Date', 'Due Date', 'Next Due Date'],
            'Machinery Location': ['Machinery Location','Machinery', 'Equipment Location', 'Machine Location'],
            'Sub Component Location': ['Sub Component Location','Component', 'Component Location', 'Part Location']
        }
        self.required_columns = list(self.column_mappings.keys())
        self.numeric_columns = ['Job Code']

    def _find_matching_column(self, columns, possible_names):
        """Find a matching column name from possible alternatives."""
        for name in possible_names:
            if name in columns:
                return name
        return None

    def _suggest_column_mappings(self, df):
        """Suggest possible column mappings for missing required columns."""
        suggestions = {}
        available_columns = df.columns.tolist()
        for required_col, alternatives in self.column_mappings.items():
            if required_col not in df.columns:
                matched_col = self._find_matching_column(available_columns, alternatives)
                if matched_col:
                    suggestions[required_col] = matched_col
        return suggestions

    def validate_columns(self, df):
        """Check if all required columns are present and suggest mappings."""
        missing_columns = []
        suggestions = self._suggest_column_mappings(df)
        for col in self.required_columns:
            if col not in df.columns and col not in suggestions:
                missing_columns.append(col)
        if missing_columns:
            error_msg = "Missing required columns:\n"
            for col in missing_columns:
                error_msg += f"- {col}\n"
            if suggestions:
                error_msg += "\nPossible column mappings found:\n"
                for req_col, matched_col in suggestions.items():
                    error_msg += f"- '{matched_col}' could be used for '{req_col}'\n"
            return False, error_msg
        return True, ""

    def validate_numeric_fields(self, df):
        """Validate numeric fields and provide detailed error messages."""
        errors = []
        for col in self.numeric_columns:
            if col not in df.columns:
                continue
            numeric_data = pd.to_numeric(df[col], errors='coerce')
            non_numeric = df[numeric_data.isna() & df[col].notna()]
            if not non_numeric.empty:
                rows = non_numeric.index.tolist()
                total_rows = len(non_numeric)
                if total_rows > 0:
                    more = f" and {total_rows - 10} more" if total_rows > 10 else ""
                    errors.append(f"Non-numeric values found in {col} at rows: {rows[:10]}{more}")
        return len(errors) == 0, "\n".join(errors)

    def validate_machinery_location(self, df):
        """Validate machinery location format with detailed error reporting."""
        if 'Machinery Location' not in df.columns:
            return False, "Missing 'Machinery Location' column"
        # Filter only engine-related entries
        engine_entries = df[df['Machinery Location'].str.contains('Main Engine|Auxiliary Engine', case=False, na=False)]
        valid_patterns = [
            # "Main Engine#1", "Main Engine - ME-C II GI#1", "Main Engine - RT - A#1", "Main Engine No1", ...
            r'^Main Engine(?:[\s-]+[\w\s\-\./]*?)?[\s-]*(?:#|No)\s*\d+$',
            r'^Main Engine[\s-]*#?\d+$',
            r'^Auxiliary Engine[\s-]*#?\d+$',
            r'^Auxiliary Engine[\s-]*No\d+$',
        ]

        import re
        pattern = '|'.join(valid_patterns)

        # Check for standard patterns first
        mask = engine_entries['Machinery Location'].apply(lambda x: bool(re.match(pattern, str(x))) if pd.notna(x) else False)

        # For entries not matching standard patterns, attempt auto-correction
        invalid_entries = engine_entries[~mask].copy()
        if not invalid_entries.empty:
            # Add a new column with corrected format (for potential fixes)
            df['_machinery_location_fixed'] = df['Machinery Location']
            # Fix only engine names ending in "No<number>", e.g. "Auxiliary EngineNo4" -> "Auxiliary Engine#4";
            # other machinery and any other "No" in a name (e.g. "Nozzle") stay as they are
            is_engine = df['Machinery Location'].str.contains('Main Engine|Auxiliary Engine', case=False, na=False)
            df.loc[is_engine, '_machinery_location_fixed'] = \
                df.loc[is_engine, 'Machinery Location'].str.replace(r'\s*No\s*(\d+)$', r'#\1', regex=True)
        
        # Original mask for validation feedback
        invalid_locations = engine_entries[~mask]
        if not invalid_locations.empty:
            sample_locations = invalid_locations['Machinery Location'].head().tolist()
            rows = invalid_locations.index.tolist()[:5]
            total_rows = len(invalid_locations)
            error_msg = (
                f"Invalid Engine Location format at rows: {rows}\n"
                f"Example invalid formats: {sample_locations}\n"
                "Expected format for engine entries:\n"
                "- Main Engine#X\n"
                "- Main Engine - MC#X\n"
                "- Main Engine - ME-C II#X\n"
                "- Auxiliary Engine#X\n"
                "where X is a number"
            )
            if total_rows > 5:
                error_msg += f"\nAnd {total_rows - 5} more rows contain invalid formats"
            return False, error_msg
        return True, ""

    def validate_data(self, df):
        """Perform all validations with comprehensive error reporting."""
        try:
            validations = [
                (self.validate_columns(df), "Column Validation"),
                (self.validate_numeric_fields(df), "Numeric Fields Validation"),
                (self.validate_machinery_location(df), "Engine Location Validation")
            ]
            errors = []
            for (is_valid, error_msg), category in validations:
                if not is_valid:
                    errors.append(f"\n{category}:\n{error_msg}")
            return len(errors) == 0, errors
        except Exception as e:
            return False, [f"Validation error: {str(e)}"]
