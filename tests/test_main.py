# tests/test_main.py
import pytest
from pathlib import Path
import pandas as pd
from tableextractor import extract_slide_data

# 测试数据配置
TEST_DIR = Path(__file__).parent
TEST_FILE = TEST_DIR / "testdata" / "cimb.pptx"

# 精确预期结果
EXPECTED_DATA = {
    "Year": ["2019", "2020", "2021", "2022", "2023"],
    "Total Revenue, US$ (m)": ["7,128", "6,190", "6,806", "7,205", "8,256"],
    "Net Income, US$ (m)": ["1,101", "284", "1,037", "1,237", "1,532"],
    "EPS ($)": ["0.1", "0.1", "0.1", "0.1", "0.2"]
}
EXPECTED_DF = pd.DataFrame(EXPECTED_DATA)

def test_company_financials():
    """验证财务数据表格的精确提取"""
    # 执行提取
    slides = extract_slide_data(TEST_FILE)
    
    # 筛选目标幻灯片
    keyword = "Company Financials"
    matched = [s for s in slides if keyword.lower() in s["title"].lower()]
    
    # 验证基础信息
    assert len(matched) == 1, f"应找到1张幻灯片，实际找到 {len(matched)} 张"
    slide = matched[0]
    
    # 验证元数据
    assert slide["slide_number"] == 10, f"预期幻灯片编号10，实际为 {slide['slide_number']}"
    assert slide["table_count"] == 1, f"应包含1个表格，实际有 {slide['table_count']} 个"
    
    # 获取实际数据
    actual_df = slide["tables"][0]
    
    # 精确比较数据
    pd.testing.assert_frame_equal(
        actual_df.sort_index(axis=1),  # 标准化列顺序
        EXPECTED_DF.sort_index(axis=1),
        check_dtype=False,
        check_exact=True
    )