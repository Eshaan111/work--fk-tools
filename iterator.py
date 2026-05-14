import os
import shutil
from pathlib import Path

def generate_lead_permutations(source_directory):
    source_root = Path(source_directory)
    output_root = source_root / "Lead_Permutations_Output"
    output_root.mkdir(exist_ok=True)
    
    # Identify subfolders (excluding the output folder)
    subfolders = [f for f in source_root.iterdir() if f.is_dir() and f.name != "Lead_Permutations_Output"]
    
    for folder in subfolders:
        valid_exts = {'.jpg', '.jpeg', '.png', '.bmp', '.webp'}
        # Original sequence 1, 2, 3, 4, 5...
        images = sorted([img for img in folder.iterdir() if img.suffix.lower() in valid_exts])
        
        if len(images) < 3:
            print(f"Skipping '{folder.name}': Not enough images to create unique leads with position 2 fixed.")
            continue
            
        # Rule 1: Image 2 (index 1) always stays at index 1
        fixed_img = images[1]
        
        # Rule 2 & 3: Find potential leads (excluding original lead and fixed image)
        # Original lead is images[0]. Fixed is images[1].
        # We only take images[2:] as new leads to ensure unique 1st place and no original sequence.
        potential_leads = images[2:] 
        
        print(f"Processing '{folder.name}': Creating {len(potential_leads)} unique lead versions.")

        for i, lead_image in enumerate(potential_leads):
            # Rule: Sequence 1 2 3 4 5 is not repeated. 
            # We build: [Lead, Fixed, Remaining...]
            
            # Get all images that are NOT the current lead and NOT the fixed image
            others = [img for img in images if img != lead_image and img != fixed_img]
            
            # Construct the new sequence
            new_sequence = [lead_image, fixed_img] + others
            
            # Create the folder
            new_folder_name = f"{folder.name}_lead_{i+2}" # Naming based on original position
            new_folder_path = output_root / new_folder_name
            new_folder_path.mkdir(exist_ok=True)
            
            # Rule: Naming of images remains 1, 2, 3, 4, 5
            for idx, img_path in enumerate(new_sequence, 1):
                new_filename = f"{idx}{img_path.suffix}"
                shutil.copy(img_path, new_folder_path / new_filename)

    print(f"\nDone! Check the folder: {output_root}")

# --- Set your directory here ---
generate_lead_permutations('C:/work-mom/HOSERY/SHORTS/CHATGPT')