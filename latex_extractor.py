import pypandoc
import re
import os

def process_word_files(directory):
    """
    Iterate through .docx files in the specified directory and convert them to LaTeX.
    """
    for file_name in os.listdir(directory):
        if file_name.endswith(".docx"):
            file_path = os.path.join(directory, file_name)
            yield file_path, os.path.splitext(file_name)[0]

def convert_word_to_latex(word_file, latex_file):
    """
    Convert a Word document to LaTeX using pandoc.
    """
    try:
        # Convert the Word file to LaTeX
        output = pypandoc.convert_file(word_file, 'latex', outputfile=latex_file)
        print(f"Successfully converted {word_file} to LaTeX")
        print(f"Save in temp file temp.tex")
    except Exception as e:
        print(f"Error during conversion: {e}")
        raise

def extract_math_expressions(latex_content):
    """
    Extract mathematical expressions from LaTeX content.
    """
    patterns = [
        r'\\\((.+?)\\\)',  # Inline math expressions \( ... \)
        r'\\\[(.+?)\\\]',  # Display math expressions \[ ... \]
        r'\$(.+?)\$',      # Inline math expressions $ ... $
        r'\\begin\{equation\}(.+?)\\end\{equation\}',  # Equation environment
        r'\\begin\{align\}(.+?)\\end\{align\}',        # Align environment
        r'\\begin\{multline\}(.+?)\\end\{multline\}',  # Multline environment
        # Add more patterns for other math environments if needed
    ]

    math_expressions = []

    for pattern in patterns:
        matches = re.findall(pattern, latex_content, re.DOTALL)
        math_expressions.extend(matches)

    return math_expressions

def read_latex_file(file_path):
    """
    Read the content of a LaTeX file.
    """
    with open(file_path, 'r', encoding='utf-8') as file:
        return file.read()

def write_math_expressions_to_file(math_expressions, output_file):
    """
    Write extracted mathematical expressions to a file.
    """
    with open(output_file, 'w', encoding='utf-8') as file:
        file.write('\\documentclass{article}\n\\usepackage{amsmath}\n\\begin{document}\n')
        for expr in math_expressions:
            file.write(expr.strip() + '\n\n')
        file.write('\\end{document}')
            
    """
    Delete the unnecessary files
    """
    # os.remove(output_file)
    # print("Deleted temporary files")

def main():
    # Directory containing the Word documents
    directory = '.\\documents'  # Replace with your directory of Word files

    # Iterate through the Word files in the directory
    for word_file_path, output_file_name in process_word_files(directory):
        latex_file_path = f'temp/{output_file_name}_bulk.tex'
        math_output_file_path = f'latex_files/{output_file_name}.tex'

        # Step 1: Convert Word to LaTeX
        convert_word_to_latex(word_file_path, latex_file_path)

        # Step 2: Read the LaTeX file
        latex_content = read_latex_file(latex_file_path)

        # Step 3: Extract mathematical expressions
        math_expressions = extract_math_expressions(latex_content)

        # Step 4: Write the extracted expressions to a file
        write_math_expressions_to_file(math_expressions, math_output_file_path)

        print(f"Extracted {len(math_expressions)} mathematical expressions from {word_file_path} to {math_output_file_path}")

if __name__ == "__main__":
    main()