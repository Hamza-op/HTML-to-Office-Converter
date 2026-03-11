import converter
import os

html_file = '2nd.html'
output_file = '2nd_test.pptx'

print("Starting conversion test...")
try:
    converter.html_to_editable_pptx(html_file, output_file)
    print("Conversion finished.")
    if os.path.exists(output_file):
        print(f"File {output_file} exists, size: {os.path.getsize(output_file)} bytes")
    else:
        print("File not found!")
except Exception as e:
    print(f"Error during conversion: {e}")
