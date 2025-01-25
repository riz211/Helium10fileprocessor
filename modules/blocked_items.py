from typing import List, Optional, Tuple
import pandas as pd
import os
from pathlib import Path
from io import BytesIO
from openpyxl.styles import Font, PatternFill

class BlockedItemsManager:
    """Manages both blocked brands and blocked Product IDs."""
    
    def __init__(self, brands_file_path: str, product_ids_file_path: str):
        """Initialize with paths for both blocked brands and Product IDs files."""
        self.brands_file_path = brands_file_path
        self.product_ids_file_path = product_ids_file_path
        self.ensure_files_exist()

    def ensure_files_exist(self) -> None:
        """Ensure both blocked items files exist with correct structure."""
        # Create brands file
        if not os.path.exists(self.brands_file_path):
            Path(os.path.dirname(self.brands_file_path)).mkdir(parents=True, exist_ok=True)
            pd.DataFrame(columns=["Blocked Brands"]).to_excel(
                self.brands_file_path,
                index=False,
                sheet_name="Blocked_Brands"
            )

        # Create Product IDs file
        if not os.path.exists(self.product_ids_file_path):
            Path(os.path.dirname(self.product_ids_file_path)).mkdir(parents=True, exist_ok=True)
            pd.DataFrame(columns=["Blocked Product IDs", "Reason"]).to_excel(
                self.product_ids_file_path,
                index=False,
                sheet_name="Blocked_Product_IDs"
            )

    def get_blocked_brands(self) -> pd.DataFrame:
        """Get current list of blocked brands."""
        try:
            df = pd.read_excel(self.brands_file_path, sheet_name="Blocked_Brands")
            if "S.No" in df.columns:
                df = df.drop(columns=["S.No"])
            df.reset_index(drop=True, inplace=True)
            df.insert(0, "S.No", range(1, len(df) + 1))
            return df
        except Exception as e:
            raise ValueError(f"Error reading blocked brands: {e}")

    def get_blocked_product_ids(self) -> pd.DataFrame:
        """Get current list of blocked Product IDs."""
        try:
            df = pd.read_excel(self.product_ids_file_path, sheet_name="Blocked_Product_IDs")
            if "S.No" in df.columns:
                df = df.drop(columns=["S.No"])
            df.reset_index(drop=True, inplace=True)
            df.insert(0, "S.No", range(1, len(df) + 1))
            return df
        except Exception as e:
            raise ValueError(f"Error reading blocked Product IDs: {e}")

    def add_brand(self, brand: str) -> Tuple[bool, str]:
        """Add a new brand to the blocked list."""
        if not brand or not brand.strip():
            return False, "Please enter a valid brand name."

        try:
            df = pd.read_excel(self.brands_file_path, sheet_name="Blocked_Brands")
            brand = brand.strip()

            if brand in df["Blocked Brands"].values:
                return False, f"The brand '{brand}' is already in the blocked list."

            new_brand_df = pd.DataFrame([{"Blocked Brands": brand}])
            df = pd.concat([df, new_brand_df], ignore_index=True)

            with pd.ExcelWriter(self.brands_file_path, engine="openpyxl", mode="w") as writer:
                df.to_excel(writer, index=False, sheet_name="Blocked_Brands")

            return True, f"Brand '{brand}' has been added to the blocked list."
        except Exception as e:
            return False, f"Error adding brand: {e}"

    def add_product_id(self, product_id: str, reason: str = "") -> Tuple[bool, str]:
        """Add a new Product ID to the blocked list."""
        if not product_id or not product_id.strip():
            return False, "Please enter a valid Product ID."

        try:
            df = pd.read_excel(self.product_ids_file_path, sheet_name="Blocked_Product_IDs")
            product_id = product_id.strip()

            if product_id in df["Blocked Product IDs"].values:
                return False, f"The Product ID '{product_id}' is already in the blocked list."

            new_product_id_df = pd.DataFrame([{
                "Blocked Product IDs": product_id,
                "Reason": reason.strip() if reason else "No reason provided"
            }])
            df = pd.concat([df, new_product_id_df], ignore_index=True)

            with pd.ExcelWriter(self.product_ids_file_path, engine="openpyxl", mode="w") as writer:
                df.to_excel(writer, index=False, sheet_name="Blocked_Product_IDs")

            return True, f"Product ID '{product_id}' has been added to the blocked list."
        except Exception as e:
            return False, f"Error adding Product ID: {e}"

    def bulk_upload_brands(self, df: pd.DataFrame) -> Tuple[bool, str]:
        """Upload multiple brands at once."""
        try:
            if "Blocked Brands" not in df.columns:
                return False, "The uploaded file must contain a 'Blocked Brands' column."

            current_df = pd.read_excel(self.brands_file_path, sheet_name="Blocked_Brands")
            if "S.No" in current_df.columns:
                current_df = current_df.drop(columns=["S.No"])

            updated_df = pd.concat([current_df, df]).drop_duplicates(
                subset=["Blocked Brands"],
                ignore_index=True
            )

            with pd.ExcelWriter(self.brands_file_path, engine="openpyxl", mode="w") as writer:
                updated_df.to_excel(writer, index=False, sheet_name="Blocked_Brands")

            return True, "Blocked Brands have been updated successfully."
        except Exception as e:
            return False, f"Error processing bulk upload: {e}"

    def bulk_upload_product_ids(self, df: pd.DataFrame) -> Tuple[bool, str]:
        """Upload multiple Product IDs at once."""
        try:
            if "Blocked Product IDs" not in df.columns:
                return False, "The uploaded file must contain a 'Blocked Product IDs' column."

            current_df = pd.read_excel(self.product_ids_file_path, sheet_name="Blocked_Product_IDs")
            if "S.No" in current_df.columns:
                current_df = current_df.drop(columns=["S.No"])

            updated_df = pd.concat([current_df, df]).drop_duplicates(
                subset=["Blocked Product IDs"],
                ignore_index=True
            )

            with pd.ExcelWriter(self.product_ids_file_path, engine="openpyxl", mode="w") as writer:
                updated_df.to_excel(writer, index=False, sheet_name="Blocked_Product_IDs")

            return True, "Blocked Product IDs have been updated successfully."
        except Exception as e:
            return False, f"Error processing bulk upload: {e}"

    def filter_data(self, df: pd.DataFrame) -> Tuple[pd.DataFrame, int, int]:
        """
        Filter dataframe to remove blocked brands and Product IDs.
        Returns: (filtered_df, blocked_brands_removed, blocked_products_removed)
        """
        initial_len = len(df)
        blocked_brands = self.get_blocked_brands()
        blocked_product_ids = self.get_blocked_product_ids()

        # Filter out blocked brands
        blocked_brands_list = [brand.upper() for brand in blocked_brands["Blocked Brands"].tolist() if brand]
        df["BRAND_UPPER"] = df["BRAND"].str.upper()
        df_after_brands = df[~df["BRAND_UPPER"].isin(blocked_brands_list)]
        df_after_brands = df_after_brands.drop(columns=["BRAND_UPPER"])

        brands_removed = initial_len - len(df_after_brands)

        # Filter out blocked Product IDs
        blocked_ids_list = [str(pid).strip() for pid in blocked_product_ids["Blocked Product IDs"].tolist() if pid]
        df_final = df_after_brands[~df_after_brands["SKU"].astype(str).isin(blocked_ids_list)]

        products_removed = len(df_after_brands) - len(df_final)

        return df_final, brands_removed, products_removed
