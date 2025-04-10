import os
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from utils import eliminate_accents


def capitalize_name(name):
    """
    Capitalize the first letter of each word in a name.
    
    Args:
        name: The name to capitalize
        
    Returns:
        The capitalized name
    """
    # Skip if name is not a string
    if not isinstance(name, str):
        return name
        
    # Split the name by spaces and capitalize each part
    words = name.split()
    capitalized_words = []
    
    for word in words:
        # Handle empty words
        if not word:
            capitalized_words.append(word)
            continue
            
        # Handle words with special characters at the beginning
        if not word[0].isalpha():
            capitalized_words.append(word)
            continue
            
        # Capitalize the first letter and keep the rest as is
        capitalized_words.append(word[0].upper() + word[1:])
    
    # Join the words back together
    return " ".join(capitalized_words)


def create_assistant_certificate(
    certificate_template,
    assistant_name,
    text_color,
    font_path,
    font_size,
    save_path='',
):
    """
    Create a certificate for an assistant of the congress.
    
    Args:
        certificate_template: Path to the certificate template image
        assistant_name: Name of the assistant to be added to the certificate
        text_color: Color of the text (hex code)
        font_path: Path to the font file
        font_size: Base font size for the name
        save_path: Directory to save the certificate
    
    Returns:
        filename: Name of the saved certificate file
    """
    # Load the font
    # Adjust font size based on name length
    adjusted_size = font_size
    if len(assistant_name) > 40:
        adjusted_size = int(font_size * 0.8)
    elif len(assistant_name) > 35:
        adjusted_size = int(font_size * 0.9)
    
    font = ImageFont.truetype(font_path, adjusted_size)
    
    # Open image and prepare drawing
    im = Image.open(certificate_template)
    width, height = im.size
    draw = ImageDraw.Draw(im)
    
    # Center the text
    text_width = draw.textlength(assistant_name, font=font)
    x = (width - text_width) / 2
    y = 350  # Vertical position for the name
    
    # Draw the name
    draw.text((x, y), assistant_name, fill=text_color, font=font)
    
    # Save certificate
    filename = f'certificado_{eliminate_accents(assistant_name)}.pdf'
    im.save(os.path.join(save_path, filename))
    return filename


# === MAIN ===
if __name__ == "__main__":
    # Configuration
    folder_path = "congreso_neurociencias"
    certificate_template = os.path.join(
        folder_path, "certificate_asistente.png"
    )
    
    # CSV with registration data
    csv_file = "Inscripción al Primer Congreso Latinoamericano de Neurociencias Cognitivas  (respuestas) - Respuestas de formulario.csv"
    csv_path = os.path.join(folder_path, csv_file)
    
    # Font settings
    font_dir = os.path.join(folder_path, "nunito-sans")
    bold_font = os.path.join(font_dir, "NunitoSans-Bold.ttf")
    
    # Output folder
    certificate_folder = os.path.join(folder_path, "certificados_asistentes")
    os.makedirs(certificate_folder, exist_ok=True)
    
    # Text color
    text_color = "#000000"
    
    # Font size
    name_font_size = 80
    
    # Process all names from the CSV
    try:
        # Read CSV file
        df = pd.read_csv(csv_path)
        
        # Column with participant names
        name_column = "Nombre y Apellido"
        
        # Counter for successful and failed certificates
        success_count = 0
        error_count = 0
        
        # Iterate through all names
        for idx, row in df.iterrows():
            # Get the name from the CSV
            original_name = row[name_column]
            
            # Check if name is valid
            if pd.isna(original_name) or not original_name:
                print(f"⚠️ Skipping empty name at row {idx+1}")
                continue
            
            # Capitalize the name
            assistant_name = capitalize_name(original_name)
                
            try:
                filename = create_assistant_certificate(
                    certificate_template=certificate_template,
                    assistant_name=assistant_name,
                    text_color=text_color,
                    font_path=bold_font,
                    font_size=name_font_size,
                    save_path=certificate_folder
                )
                print(f"✅ Certificate created: {filename}")
                success_count += 1
            except Exception as e:
                err_msg = f"❌ Error creating certificate for {assistant_name}"
                print(f"{err_msg}: {str(e)}")
                error_count += 1
        
        # Print summary
        print("\n=== SUMMARY ===")
        print(f"Total certificates created: {success_count}")
        print(f"Failed certificates: {error_count}")
        print(f"Total processed: {success_count + error_count}")
        
    except Exception as e:
        print(f"❌ Error processing CSV file: {str(e)}") 