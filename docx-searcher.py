import os
import docx
import argparse

def search_keywords_in_docx_insensitive(file_path, keywords):
    doc = docx.Document(file_path)
    #text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
    for i, paragraph in enumerate(doc.paragraphs):
        for keyword in keywords:
            if keyword.lower() in paragraph.text.lower():
                print("found some insensitive materials in " + file_path)
                k = 1
                while not doc.paragraphs[i-k].style.name.startswith('Heading'):
                    print("some other paragraph: " + doc.paragraphs[i-k].text)
                    k += 1
                print("Found in heading: " + doc.paragraphs[i-k].text)

def search_keywords_in_docx(file_path, keywords):
    doc = docx.Document(file_path)
    text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
    for keyword in keywords:
        if keyword in text:
            print("Found some materials")
            return True
    return False

def process_directory(directory_path, keywords, case_insensitive):
    for root, _, files in os.walk(directory_path):
        for file_name in files:
            if file_name.endswith('.docx') and not file_name.startswith('~'):
                file_path = os.path.join(root, file_name)
                if (case_insensitive):
                    print("SO INSENSITIVE")
                    search_keywords_in_docx_insensitive(file_path, keywords)
                        #print(f"Found keywords in: {file_path}")
                else: 
                    search_keywords_in_docx(file_path, keywords)
                        #print(f"Found keywords in: {file_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Search for keywords in DOCX files in a directory.")
    parser.add_argument("directory_path", help="Path to the directory containing DOCX files.")
    parser.add_argument("keywords", nargs="+", help="Keywords to search for in the DOCX files.")
    parser.add_argument("-i", "--case_insensitive", action="store_true", help="Search case insensitive")
    args = parser.parse_args()

    process_directory(args.directory_path, args.keywords, args.case_insensitive)
