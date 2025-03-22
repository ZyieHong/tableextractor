# src/tableextractor/cli.py
import click
from pathlib import Path
from .main import extract_slide_data

@click.command()
@click.argument("filepath", type=click.Path(exists=True))
@click.argument("keyword")
def tableextractor(filepath, keyword):
    """根据关键词提取PPT表格数据"""
    try:
        # 基本文件验证
        if not filepath.endswith(".pptx"):
            raise ValueError("只支持 .pptx 文件")
            
        # 核心处理流程
        slides = extract_slide_data(filepath)
        results = [s for s in slides if keyword.lower() in s["title"].lower()]
        
        # 结果输出
        if not results:
            print("没有找到匹配的幻灯片")
            return
            
        print(f"找到 {len(results)} 张匹配的幻灯片")
        
        for slide in results:
            print("\n" + "=" * 80)
            print(f"幻灯片 #{slide['slide_number']}")
            print(f"标题: {slide['title']}")
            
            if not slide["tables"]:
                print("本页无表格")
                continue
                
            for idx, df in enumerate(slide["tables"], 1):
                print(f"\n表格 {idx}:")
                print(df.to_string(index=False))
                
    except Exception as e:
        print(f"发生错误: {str(e)}")

if __name__ == "__main__":
    tableextractor()