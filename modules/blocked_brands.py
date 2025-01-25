from typing import List, Optional, Tuple
import pandas as pd
import os
from pathlib import Path
from io import BytesIO
from openpyxl.styles import Font, PatternFill

class BlockedBrandsManager:
    def __init__(self, file_path: str):
        """Initialize BlockedBrandsManager with file path."""
        self.file_path = file_path
        self.ensure_file_exists()

    def ensure_file_exists(self) -> None:
        """Ensure the blocked brands file exists with correct structure."""
        if not os.path.exists(self.file_path):
            # Create directory if it doesn't exist
            Path(os.path.dirname(self.file_path)).mkdir(parents=True, exist_ok=True)

            # Create new file with correct structure
            pd.DataFrame(columns=["Blocked Brands"]).to_excel(
                self.file_path,
                index=False,
                sheet_name="Blocked_Brands"
            )

    def get_blocked_brands(self) -> pd.DataFrame:
        """Get current list of blocked brands."""
        try:
            df = pd.read_excel(self.file_path, sheet_name="Blocked_Brands")

            # Remove S.No if it exists, then add it fresh
            if "S.No" in df.columns:
                df = df.drop(columns=["S.No"])

            # Reset index and add S.No
            df.reset_index(drop=True, inplace=True)
            df.insert(0, "S.No", range(1, len(df) + 1))
            return df

        except Exception as e:
            raise ValueError(f"Error reading blocked brands: {e}")

    def add_brand(self, brand: str) -> Tuple[bool, str]:
        """Add a new brand to the blocked list."""
        if not brand or not brand.strip():
            return False, "Please enter a valid brand name."

        try:
            df = pd.read_excel(self.file_path, sheet_name="Blocked_Brands")
            brand = brand.strip()

            if "Blocked Brands" not in df.columns:
                df = pd.DataFrame(columns=["Blocked Brands"])

            if brand in df["Blocked Brands"].values:
                return False, f"The brand '{brand}' is already in the blocked list."

            new_brand_df = pd.DataFrame([{"Blocked Brands": brand}])
            df = pd.concat([df, new_brand_df], ignore_index=True)

            with pd.ExcelWriter(self.file_path, engine="openpyxl", mode="w") as writer:
                df.to_excel(writer, index=False, sheet_name="Blocked_Brands")

            return True, f"Brand '{brand}' has been added to the blocked list."

        except Exception as e:
            return False, f"Error adding brand: {e}"

    def bulk_upload(self, df: pd.DataFrame) -> Tuple[bool, str]:
        """Upload multiple brands at once."""
        try:
            if "Blocked Brands" not in df.columns:
                return False, "The uploaded file must contain a 'Blocked Brands' column."

            current_df = pd.read_excel(self.file_path, sheet_name="Blocked_Brands")

            if "S.No" in current_df.columns:
                current_df = current_df.drop(columns=["S.No"])

            updated_df = pd.concat([current_df, df]).drop_duplicates(
                subset=["Blocked Brands"],
                ignore_index=True
            )

            with pd.ExcelWriter(self.file_path, engine="openpyxl", mode="w") as writer:
                updated_df.to_excel(writer, index=False, sheet_name="Blocked_Brands")

            return True, "Blocked Brands have been updated successfully."

        except Exception as e:
            return False, f"Error processing bulk upload: {e}"

    def export_blocked_brands(self) -> Tuple[BytesIO, Optional[str]]:
        """Export blocked brands to an Excel file with proper formatting."""
        try:
            df = self.get_blocked_brands()

            # Create a buffer for the Excel file
            buffer = BytesIO()

            # Create Excel writer
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                # Save only the Blocked Brands column when exporting
                export_df = pd.DataFrame(df["Blocked Brands"])
                export_df.to_excel(writer, index=False, sheet_name="Blocked_Brands")

                # Get the worksheet to apply formatting
                worksheet = writer.sheets["Blocked_Brands"]

                # Format header
                for cell in worksheet[1]:
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")

                # Adjust column width
                worksheet.column_dimensions['A'].width = 30

            buffer.seek(0)
            return buffer, None

        except Exception as e:
            return None, f"Error exporting blocked brands: {str(e)}"