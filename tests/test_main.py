import pytest
from pathlib import Path
import pandas as pd
from tableextractor import extract_slide_data

# test data configuration
test_dir = Path(__file__).parent
test_file = test_dir / "test_data" / "cimb.pptx"

expected_data = {
    "Year": ["2019", "2020", "2021", "2022", "2023"],
    "Total Revenue, US$ (m)": ["7,128", "6,190", "6,806", "7,205", "8,256"],
    "Net Income, US$ (m)": ["1,101", "284", "1,037", "1,237", "1,532"],
    "EPS ($)": ["0.1", "0.1", "0.1", "0.1", "0.2"]
}
expected_df = pd.DataFrame(expected_data)

def test_company_financials():
    """Verify precise extraction of financial data tables"""
    # Perform extraction
    slides = extract_slide_data(test_file)
    
    keyword = "Company Financials"
    matched = [s for s in slides if keyword.lower() in s["title"].lower()]
    
    # Verify basic information
    assert len(matched) == 1, f"Expected 1 slide, but found {len(matched)}"
    slide = matched[0]
    
    # Verify metadata
    assert slide["slide_number"] == 10, f"Expected slide number 10, but got {slide['slide_number']}"
    assert slide["table_count"] == 1, f"Expected 1 table, but found {slide['table_count']}"
    
    # Get actual data
    actual_df = slide["tables"][0]
    
    # Precisely compare data
    pd.testing.assert_frame_equal(
        actual_df.sort_index(axis=1),  # Standardize column order
        expected_df.sort_index(axis=1),
        check_dtype=False,
        check_exact=True
    )
    
if __name__ == "__main__":
    test_company_financials()