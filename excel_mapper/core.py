import pandas as pd
import re
import os
from typing import List, Optional

class ExcelMapper:
    """
    A class to read and manipulate Excel files with dynamic attribute access.
    
    This class reads an Excel file where the first row contains column headers
    and subsequent rows contain data. Column names are automatically sanitized
    and converted to attribute names for easy access.
    
    Attributes:
        file_path (str): Path to the Excel file
        rows (List[Row]): List of Row objects containing the data
        column_mapping (dict): Mapping of sanitized column names to original names
    
    Example:
        >>> mapper = ExcelMapper("financial_data.xlsx")
        >>> # Access row 2 data (index 1 in 0-based indexing)
        >>> transaction_id = mapper.rows[1].transaction_id  # For "Transaction ID" column
        >>> amount = mapper.rows[1].amount_usd  # For "Amount (USD)" column
    
    Note:
        - Column names are sanitized: special characters are replaced with underscores
        - Duplicate column names get suffixes: `{colname}_A`, `{colname}_B`, etc.
        - Rows are 0-indexed: mapper.rows[0] corresponds to Excel row 2
        - NaN values are converted to None for better Python compatibility
    """

    def __init__(self, file_path: str):
        """
        Initialize ExcelMapper with file path and read the Excel file.
        
        Args:
            file_path (str): Path to the Excel file
            
        Example:
            >>> mapper = ExcelMapper("financial_report.xlsx")
            >>> print(f"Loaded {len(mapper.rows)} transactions")
        """
        self.file_path = file_path
        self.rows: List[Row] = []
        self._read_excel()
    
    def _read_excel(self):
        """Read the Excel file and process the data."""
        try:
            # Read the Excel file
            df = pd.read_excel(self.file_path)
            
            # Get column names and create sanitized versions
            original_columns = df.columns.tolist()
            sanitized_columns = self._sanitize_column_names(original_columns)
            
            # Create a mapping of sanitized column names to original column names
            self.column_mapping = dict(zip(sanitized_columns, original_columns))
            
            # Convert DataFrame to list of dictionaries
            data = df.to_dict('records')
            
            # Create Row objects for each data row
            for row_data in data:
                row_obj = Row(row_data, sanitized_columns, self.column_mapping)
                self.rows.append(row_obj)
                
        except Exception as e:
            raise Exception(f"Error reading Excel file: {str(e)}")
    
    def _sanitize_column_names(self, column_names: List[str]) -> List[str]:
        """
        Sanitize column names according to the specified rules.
        
        Args:
            column_names (List[str]): List of original column names
            
        Returns:
            List[str]: List of sanitized column names
            
        Example:
            >>> mapper._sanitize_column_names(["Transaction ID", "Amount (USD)"])
            ['transaction_id', 'amount_usd']
        """
        sanitized_names = []
        name_count = {}
        
        for i, col_name in enumerate(column_names):
            # Convert to string if not already
            col_name_str = str(col_name)
            
            # Convert to lowercase and replace special characters with underscores
            sanitized = col_name_str.lower()
            sanitized = re.sub(r'[^a-z0-9]+', '_', sanitized)
            sanitized = re.sub(r'_+', '_', sanitized)
            sanitized = sanitized.strip('_')
            
            # Handle duplicate column names
            if sanitized in name_count:
                name_count[sanitized] += 1
                # Get column letter (A, B, C, ...)
                col_letter = chr(65 + i)  # 65 is ASCII for 'A'
                sanitized = f"{sanitized}_{col_letter}"
            else:
                name_count[sanitized] = 1
            
            sanitized_names.append(sanitized)
        
        return sanitized_names

    def get_column_mapping(self) -> dict:
        """
        Get a mapping of original column names to sanitized attribute names.
        
        Returns:
            dict: Dictionary where keys are original column names and values 
                are corresponding sanitized attribute names
                
        Example:
            >>> mapper.get_column_mapping()
            {
                'Transaction ID': 'transaction_id',
                'Amount (USD)': 'amount_usd',
                'Transaction Date': 'transaction_date'
            }
        """
        return {original: sanitized for sanitized, original in self.column_mapping.items()}

    def get_original_columns(self) -> List[str]:
        """
        Get the list of original column names from the Excel file.
        
        Returns:
            List[str]: List of original column names in their original order
            
        Example:
            >>> mapper.get_original_columns()
            ['Transaction ID', 'Amount (USD)', 'Transaction Date', 'Category']
        """
        return list(self.column_mapping.values())

    def get_attribute_names(self) -> List[str]:
        """
        Get the list of sanitized attribute names for dynamic access.
        
        Returns:
            List[str]: List of sanitized attribute names in column order
            
        Example:
            >>> mapper.get_attribute_names()
            ['transaction_id', 'amount_usd', 'transaction_date', 'category']
        """
        return list(self.column_mapping.keys())

    def get_column_info(self) -> List[dict]:
        """
        Get detailed information about all columns including both original and attribute names.
        
        Returns:
            List[dict]: List of dictionaries with column information, each containing:
                - 'original_name': Original column name from Excel
                - 'attribute_name': Sanitized attribute name for access
                - 'index': Column position (0-based)
                
        Example:
            >>> mapper.get_column_info()
            [
                {'index': 0, 'original_name': 'Transaction ID', 'attribute_name': 'transaction_id'},
                {'index': 1, 'original_name': 'Amount (USD)', 'attribute_name': 'amount_usd'},
                {'index': 2, 'original_name': 'Transaction Date', 'attribute_name': 'transaction_date'}
            ]
        """
        return [
            {
                'index': i,
                'original_name': original_name,
                'attribute_name': attr_name
            }
            for i, (attr_name, original_name) in enumerate(self.column_mapping.items())
        ]
    
    def save_excel(self, file_path: str = None, overwrite: bool = False) -> None:
        """
        Save the current data back to an Excel file.
        
        Args:
            file_path (str, optional): Path to save the file. If None, overwrites original file.
            overwrite (bool): If True, allows overwriting existing file. Default False.
            
        Raises:
            ValueError: If file_path is not provided and overwrite is False
            FileExistsError: If file exists and overwrite is False
            
        Example:
            >>> # Update values and save
            >>> mapper.rows[0].transaction_id = "TXN_2024_001"
            >>> mapper.rows[0].amount_usd = 2500.75
            >>> mapper.save_excel("updated_financials.xlsx")  # Save to new file
            >>> mapper.save_excel(overwrite=True)             # Overwrite original file
        """
        if file_path is None:
            if not overwrite:
                raise ValueError("overwrite must be True to overwrite original file")
            file_path = self.file_path
        
        if not overwrite and os.path.exists(file_path):
            raise FileExistsError(f"File {file_path} already exists. Use overwrite=True to overwrite.")
        
        # Convert rows back to DataFrame
        data = []
        for row in self.rows:
            row_dict = {}
            for attr_name, orig_col in self.column_mapping.items():
                value = getattr(row, attr_name, None)
                row_dict[orig_col] = value
            data.append(row_dict)
        
        # Create DataFrame and save
        df = pd.DataFrame(data)
        df.to_excel(file_path, index=False)
        if overwrite:
            print(f"== Updated the sheet ==")
        else:
            print(f"== New excel file created ==")


    def update_row(self, row_index: int, **kwargs) -> None:
        """
        Update specific values in a row using attribute names.
        
        Args:
            row_index (int): Index of the row to update (0-based)
            **kwargs: Attribute name and value pairs to update
            
        Raises:
            IndexError: If row_index is out of range
            AttributeError: If attribute name doesn't exist
            
        Example:
            >>> mapper.update_row(0, transaction_id="TXN_2024_001", amount_usd=1500.50)
        """
        if row_index < 0 or row_index >= len(self.rows):
            raise IndexError(f"Row index {row_index} out of range")
        
        row = self.rows[row_index]
        for attr_name, value in kwargs.items():
            if not hasattr(row, attr_name):
                raise AttributeError(f"Attribute '{attr_name}' does not exist")
            setattr(row, attr_name, value)


    def update_column(self, column_attr: str, new_values: List) -> None:
        """
        Update an entire column with new values.
        
        Args:
            column_attr (str): Sanitized attribute name of the column
            new_values (List): List of new values for the column
            
        Raises:
            AttributeError: If column attribute doesn't exist
            ValueError: If new_values length doesn't match row count
            
        Example:
            >>> mapper.update_column("amount_usd", [100.50, 200.75, 300.25])
        """
        if not hasattr(self.rows[0], column_attr):
            raise AttributeError(f"Column attribute '{column_attr}' does not exist")
        
        if len(new_values) != len(self.rows):
            raise ValueError(f"Expected {len(self.rows)} values, got {len(new_values)}")
        
        for i, row in enumerate(self.rows):
            setattr(row, column_attr, new_values[i])


    def add_row(self, **kwargs) -> None:
        """
        Add a new row to the Excel data.
        
        Args:
            **kwargs: Attribute name and value pairs for the new row
            
        Example:
            >>> mapper.add_row(transaction_id="TXN_2024_999", amount_usd=999.99, category="Office Supplies")
        """
        # Create empty data template
        empty_data = {orig_col: None for orig_col in self.column_mapping.values()}
        
        # Update with provided values
        for attr_name, value in kwargs.items():
            if attr_name not in empty_data:
                raise AttributeError(f"Attribute '{attr_name}' does not exist")
            empty_data[attr_name] = value
        
        # Create new row
        new_row = Row(empty_data, list(self.column_mapping.keys()), self.column_mapping)
        self.rows.append(new_row)
    
    def __getitem__(self, index: int):
        """Allow indexing to access rows directly."""
        return self.rows[index]
    
    def __len__(self):
        """Return the number of rows."""
        return len(self.rows)
    
    def __iter__(self):
        """Make the object iterable."""
        return iter(self.rows)


class Row:
    def __init__(self, data: dict, sanitized_columns: List[str], column_mapping: dict):
        """
        Initialize a Row object with data and column information.
        
        Example:
            >>> row = mapper.rows[0]
            >>> print(row.transaction_id)  # Access transaction ID
            >>> row.amount_usd = 1500.75  # Update amount
        """
        self._data = data
        self._sanitized_columns = sanitized_columns
        self._column_mapping = column_mapping
        self._original_columns = list(data.keys())
        
        # Create dynamic attributes for original data
        for orig_col, sanitized_col in zip(self._original_columns, sanitized_columns):
            value = data[orig_col]
            if pd.isna(value):
                value = None
            setattr(self, sanitized_col, value)
    
    def __getattr__(self, name: str):
        """Handle attribute access for non-existent attributes."""
        raise AttributeError(f"'{self.__class__.__name__}' object has no attribute '{name}'")
    
    def to_dict(self) -> dict:
        """Return the row data as a dictionary including dynamically added attributes."""
        result = self._data.copy()
        
        # Add dynamically added attributes that aren't in original data
        all_attrs = [attr for attr in dir(self) if not attr.startswith('_') and not callable(getattr(self, attr))]
        original_attrs = list(self._data.keys())
        
        for attr in all_attrs:
            if attr not in original_attrs:
                result[attr] = getattr(self, attr)
        
        return result
    
    def __repr__(self) -> str:
        """String representation of the Row object including dynamically added attributes."""
        attrs = []
        # Get all non-private, non-callable attributes
        for attr in dir(self):
            if not attr.startswith('_') and not callable(getattr(self, attr)):
                value = getattr(self, attr)
                attrs.append(f"{attr}={repr(value)}")
        
        return f"Row({', '.join(attrs)})"
    
    def get_original_column_name(self, sanitized_name: str) -> Optional[str]:
        """Get the original column name from a sanitized name."""
        return self._column_mapping.get(sanitized_name)
