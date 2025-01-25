import pandas as pd
import re
from typing import List, Optional, Tuple
from .excel_utils import calculate_shipping_cost
from .blocked_brands import BlockedBrandsManager

def extract_weight_with_packs(title: str) -> Optional[float]:
    """Extract weight from title, considering pack sizes."""
    try:
        # Extract weight
        weight_match = re.search(
            r"(\d+(\.\d+)?)\s*(?:oz|ounces|ounce|fl\. oz\.|fluid ounce|fl oz|fluid ounces)",
            title,
            re.IGNORECASE
        )
        if not weight_match:
            return None

        single_weight = float(weight_match.group(1))

        # Extract pack size
        pack_match = re.search(r"(?:\b(\d+)\s*pack\b|\bpack of\s*(\d+))", title, re.IGNORECASE)
        pack_size = int(pack_match.group(1) or pack_match.group(2)) if pack_match else 1

        # Add packaging weight
        if "fl oz" in weight_match.group(0).lower():
            single_weight += 10  # Fluid container weight
        else:
            single_weight += 6   # Regular packaging weight

        # Calculate total weight in pounds
        total_weight = (single_weight * pack_size) / 16
        return round(total_weight, 2)

    except Exception:
        return None

def process_dataframes(
    dfs: List[pd.DataFrame],
    blocked_brands_manager: BlockedBrandsManager = None,
    shipping_legend: Optional[pd.DataFrame] = None
) -> Tuple[pd.DataFrame, int]:  # Returns (DataFrame, number of blocked items removed)
    """Process multiple dataframes into a single formatted dataframe."""
    if not dfs:
        return pd.DataFrame(), 0

    # Combine dataframes
    df = pd.concat(dfs, ignore_index=True)
    removed_count = 0

    # First standardize and rename columns
    column_mapping = {
        "Product Details": "TITLE",
        "Brand": "BRAND",
        "Product ID": "SKU",
        "UPC Code": "UPC/ISBN",
        "Price": "COST_PRICE"
    }
    df.rename(columns=column_mapping, inplace=True)

    # Then filter out blocked brands if manager is provided
    if blocked_brands_manager is not None:
        try:
            blocked_brands = blocked_brands_manager.get_blocked_brands()
            blocked_list = [brand.upper() for brand in blocked_brands["Blocked Brands"].tolist() if brand]

            # Store initial length for reporting
            initial_len = len(df)

            # Convert brands to uppercase for comparison and filter
            df["BRAND_UPPER"] = df["BRAND"].str.upper()
            df = df[~df["BRAND_UPPER"].isin(blocked_list)]
            df = df.drop(columns=["BRAND_UPPER"])  # Remove temporary column

            removed_count = initial_len - len(df)
        except Exception as e:
            print(f"Error filtering blocked brands: {e}")

    # Format SKU
    if "SKU" in df.columns:
        df["SKU"] = df["SKU"].astype(str).str.replace(",", "").str.strip()

    # Clean TITLE
    if "TITLE" in df.columns:
        df["TITLE"] = (df["TITLE"]
            .str.replace(r"\(W\+\)", "", regex=True)
            .str.replace(r"\(SP\)", "", regex=True)
            .str.replace(r"\(P\)", "", regex=True)
            .str.strip())

    # Format UPC/ISBN
    if "UPC/ISBN" in df.columns:
        df["UPC/ISBN"] = (df["UPC/ISBN"]
            .apply(lambda x: str(int(float(x))) if pd.notnull(x) else "")
            .str.zfill(12))

    # Format COST_PRICE
    if "COST_PRICE" in df.columns:
        df["COST_PRICE"] = (df["COST_PRICE"]
            .astype(str)
            .str.replace(r"[$,]", "", regex=True)
            .astype(float)
            .round(2))

    # Add standard columns
    df["HANDLING COST"] = 0.75
    df["QUANTITY"] = 1
    df["ITEM LOCATION"] = "WALMART"

    # Calculate weights
    df["ITEM WEIGHT (pounds)"] = df["TITLE"].apply(extract_weight_with_packs)

    # Add SHIPPING COST column based on weight and shipping legend
    if shipping_legend is not None:
        df["SHIPPING COST"] = df["ITEM WEIGHT (pounds)"].apply(
            lambda weight: calculate_shipping_cost(weight, shipping_legend)
        )

    # Add RETAIL PRICE column
    if all(col in df.columns for col in ["COST_PRICE", "SHIPPING COST", "HANDLING COST"]):
        df["RETAIL PRICE"] = df.apply(
            lambda row: round(
                (row["COST_PRICE"] + row["SHIPPING COST"] + row["HANDLING COST"]) * 1.35, 2
            ) if not (pd.isnull(row["COST_PRICE"]) or pd.isnull(row["SHIPPING COST"]) or pd.isnull(row["HANDLING COST"])) else None,
            axis=1
        )

    # Add MIN PRICE and MAX PRICE columns
    if all(col in df.columns for col in ["SHIPPING COST", "ITEM WEIGHT (pounds)", "RETAIL PRICE"]):
        df["MIN PRICE"] = df["RETAIL PRICE"]
        df["MAX PRICE"] = df["RETAIL PRICE"].apply(lambda x: round(x * 1.35, 2) if x is not None else None)

    # Remove duplicates
    df.drop_duplicates(inplace=True)

    # Sort rows with missing weights to the end
    df['Missing Weight'] = df['ITEM WEIGHT (pounds)'].isnull()
    df = df.sort_values(by='Missing Weight', ascending=True)
    df = df.drop(columns=['Missing Weight'])

    return df, removed_count