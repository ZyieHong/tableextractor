# src/tableextractor/cli.py
import click
from pathlib import Path
from .main import extract_slide_data

@click.command()
@click.argument("filepath", type=click.Path(exists=True))
@click.argument("keyword")
def tableextractor(filepath, keyword): # Extract PPT table data based on a keyword
    try:
        # Basic file validation
        if not filepath.endswith(".pptx"):
            raise ValueError("Only .pptx files are supported")
            
        # Core processing workflow
        slides = extract_slide_data(filepath)
        results = [s for s in slides if keyword.lower() in s["title"].lower()]
        
        # Result output
        if not results:
            print("No matching slides found")
            return
            
        print(f"Found {len(results)} matching slides")
        
        for slide in results:
            print("\n" + "=" * 80)
            print(f"Slide #{slide['slide_number']}")
            print(f"Title: {slide['title']}")
            
            if not slide["tables"]:
                print("No tables in this page")
                continue
                
            for idx, df in enumerate(slide["tables"], 1):
                print(f"\nTable {idx}:")
                print(df.to_string(index=False))
                
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    tableextractor()