import os
import zipfile

def create_oxt(output_path, source_dir):
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as oxt:
        for root, dirs, files in os.walk(source_dir):
            for file in files:
                if file.endswith('.oxt') or file == 'package.py':
                    continue
                file_path = os.path.join(root, file)
                arc_name = os.path.relpath(file_path, source_dir)
                oxt.write(file_path, arc_name)
    print(f"Extension packaged at: {output_path}")

if __name__ == "__main__":
    source = os.getcwd()
    OXT_NAME = "Lexico.oxt"
    output = os.path.join(source, OXT_NAME)
    create_oxt(output, source)
