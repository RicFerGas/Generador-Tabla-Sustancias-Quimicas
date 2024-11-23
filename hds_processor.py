# hds_processor.py
from pathlib import Path
import os
import sys
import json
from typing import Callable, List, Tuple, Optional
from preprocess import DocumentPreprocessor
from extract_info import extract_info_from_hds_txt
from excel_postprocess import GeneradorTablaSustQ

class HDSProcessor:
    """
    Process multiple HDS documents from a folder, generate structured data,
    and export to Excel and JSON formats.
    """
    def __init__(self, client, config_path: str = "data_sets"):
        self.client = client
        if getattr(sys, 'frozen', False):
            # Running as exe
            base_path = sys._MEIPASS
        else:
            # Running as script
            base_path = os.path.dirname(os.path.abspath(__file__))
            
        self.config_path = os.path.join(base_path, 'data_sets')
        print(f"Using config path: {self.config_path}")
        
        self.doc_processor = DocumentPreprocessor()
        self.excel_generator = GeneradorTablaSustQ(config_path)
        
    def process_folder(self,
                      folder_path: str,
                      project_name: str,
                      progress_callback: Optional[Callable] = None) -> Tuple[List[str], List[dict]]:
        """
        Process all supported documents in a folder.

        Args:
            folder_path: Path to folder containing HDS documents
            project_name: Name of the project
            excel_output: Path for Excel output file
            json_output: Path for JSON output file
            progress_callback: Optional callback function for progress updates

        Returns:
            Tuple containing list of errors and list of processed data
        """
        # Output paths
        json_output = f"{project_name}_raw_data.json"
        excel_output = f"{project_name}_Tabla_de_HDSs.xlsx"
        folder = Path(folder_path)
        all_data = []
        errors = []
        
        # Get all supported files
        files = [f for f in folder.glob('*') 
                if f.suffix.lower() in self.doc_processor.supported_extensions]
        total_files = len(files)

        for index, file_path in enumerate(files):
            try:
                # Update progress
                if progress_callback:
                    progress = (index * 100) // total_files
                    progress_callback(progress, file_path.name)

                # Extract text
                text = self.doc_processor.extract_text(str(file_path))
                if not text or text in ["Unsupported file format", "Invalid PDF file", "Text extraction failed"]:
                    raise ValueError(f"Text extraction failed for {file_path}")

                # Process through LLM
                hds_data = extract_info_from_hds_txt(text, self.client)
                
                # Convert to dict and add filename
                hds_dict = hds_data.model_dump()
                hds_dict["Archivo"] = file_path.name
                
                all_data.append(hds_dict)
                print(f"Successfully processed: {file_path.name}")

            except Exception as e:
                print(f"Error processing {file_path.name}: {str(e)}")
                errors.append(str(file_path))

        # Save JSON output
        with open(json_output, 'w', encoding='utf-8') as f:
            json.dump(all_data, f, indent=2, ensure_ascii=False)

        # Generate Excel with template
        try:
            flattened_data = self.excel_generator.flatten_hds_data(all_data)
            self.excel_generator.export_to_excel_with_template(
                flattened_data,
                excel_output
            )
        except Exception as e:
            print(f"Error generating Excel: {str(e)}")
            errors.append("Excel generation failed")

        # Final progress update
        if progress_callback:
            progress_callback(100, "Complete")

        return errors, all_data
if __name__ == "__main__":
    # Example usage
    import os
    import openai
    from dotenv import load_dotenv

    load_dotenv()
    openai.api_key = os.getenv("OPENAI_API_KEY")
    client = openai.OpenAI()
    hds_processor = HDSProcessor(client)
    errors, all_data = hds_processor.process_folder(
        "ejemplo",
        "ejemplo/outputs/ejemplo_gral_run")
    
    print(f"Errors: {errors}")
    print(f"Processed data: {len(all_data)}")