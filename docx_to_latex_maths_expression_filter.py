import pypandoc
import re
import os

def convert_word_to_latex(word_file, latex_file):
    """
    Convert a Word document to LaTeX using pandoc.
    """
    try:
        # Convert the Word file to LaTeX
        output = pypandoc.convert_file(word_file, 'latex', outputfile=latex_file)
        print(f"Successfully converted {word_file} to {latex_file}")
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
        for expr in math_expressions:
            file.write(expr.strip() + '\n\n')

def main():
    # Path to the Word document
    word_file_path = './c.docx'  # Replace with your Word file path
    # Path to the output LaTeX file
    latex_file_path = 'output.tex'
    # Path to the output file for extracted math expressions
    math_output_file_path = 'math_expressions.txt'

    # Step 1: Convert Word to LaTeX
    convert_word_to_latex(word_file_path, latex_file_path)

    # Step 2: Read the LaTeX file
    latex_content = read_latex_file(latex_file_path)

    # Step 3: Extract mathematical expressions
    math_expressions = extract_math_expressions(latex_content)

    # Step 4: Write the extracted expressions to a file
    write_math_expressions_to_file(math_expressions, math_output_file_path)

    print(f"Extracted {len(math_expressions)} mathematical expressions to {math_output_file_path}")

if __name__ == "__main__":
    main()