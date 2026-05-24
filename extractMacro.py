from oletools.olevba import VBA_Parser

def extract_excel_macros(file_path):
    # Initialize the VBA Parser with your Excel file
    vba_parser = VBA_Parser(file_path)
    
    # Check if the file actually contains macros
    if vba_parser.detect_vba_macros():
        print(f"Macros detected in {file_path}\n")
        
        # Iterate through all extracted macro modules
        for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
            if vba_filename[-4:] == ".cls":
                continue
            folderName = file_path.partition("\\")[0]
            with open(f'{folderName}\\{vba_filename}', 'w') as f:
                normalized_code = vba_code.replace('\r\n', '\n')
                f.write(normalized_code)
            print(f"{file_path} - Module: {vba_filename}")
    else:
        print("No macros found in this file.")

# Usage
extract_excel_macros('vba Library\\vba Library.xlsm')
extract_excel_macros('ImportExport\\ImportExport.xlsm')
extract_excel_macros('TraceFormulaTool\\TraceFormulaTool.xlsm')