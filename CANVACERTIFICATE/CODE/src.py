from PIL import Image, ImageDraw, ImageFont
import pandas as pd

# Load the Excel file
excel_file = "attende.xlsx"
data = pd.read_excel(excel_file, sheet_name='Sheet 1')

# Font setup
font_path = "martinithai-neue-slab-regular.otf"  # Adjust this path
initial_font_size = 130  # Start with a large font size

# Define a function to generate certificates
def generate_certificates(students, template_path, year_label, max_width):
    for student in students:
        name = student['Name']
        # Load the appropriate template
        template = Image.open(template_path)
        draw = ImageDraw.Draw(template)
        
        # Start with the initial font size
        font_size = initial_font_size
        font = ImageFont.truetype(font_path, font_size)
        
        # Reduce font size until the text fits within the max width
        name_bbox = draw.textbbox((0, 0), name, font=font)
        text_width = name_bbox[2] - name_bbox[0]
        text_height = name_bbox[3] - name_bbox[1]

        while text_width > max_width and font_size > 10:  # Don't reduce font size too much
            font_size -= 1
            font = ImageFont.truetype(font_path, font_size)
            name_bbox = draw.textbbox((0, 0), name, font=font)
            text_width = name_bbox[2] - name_bbox[0]
            text_height = name_bbox[3] - name_bbox[1]

        # Calculate the position to center the text
        image_width, image_height = template.size
        name_x = (image_width - text_width) / 4  # Adjust horizontal positioning as needed
        name_y = 660  # Adjust as necessary for your design

        # Draw the text on the certificate
        draw.text((name_x, name_y), name, font=font, fill="#5cb41f")

        # Save the certificate
        certificate_filename = f"certificate_{name.replace(' ', '_')}_{year_label}.png"
        template.save(certificate_filename)

# Set the maximum width that the text should not exceed
max_text_width = 1000  # Adjust this based on your template

# Filter students by academic year and generate certificates
second_year_students = data[data['Year '] == 2]
third_year_students = data[data['Year '] == 3]
fourth_year_students = data[data['Year '] == 4]

generate_certificates(second_year_students.to_dict(orient='records'), "certificate2.png", "2nd Year", max_text_width)
generate_certificates(third_year_students.to_dict(orient='records'), "certificate3.png", "3rd Year", max_text_width)
generate_certificates(fourth_year_students.to_dict(orient='records'), "certificate4.png", "4th Year", max_text_width)
