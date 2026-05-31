import os
from io import BytesIO
from PIL import Image
from django.conf import settings

def generate_image_pdf(document):
    image_paths = []

    # Add main image
    if document.image:
        image_paths.append(os.path.join(settings.MEDIA_ROOT, document.image.name))

    # Add extra pages
    for page in document.pages.all():
        if page.image:
            image_paths.append(os.path.join(settings.MEDIA_ROOT, page.image.name))

    images = []
    for path in image_paths:
        try:
            img = Image.open(path).convert('RGB')
            images.append(img)
        except Exception as e:
            print(f"Error loading image: {path} – {e}")

    if not images:
        raise ValueError("No images found.")

    pdf_stream = BytesIO()
    first_img = images[0]
    other_imgs = images[1:]
    first_img.save(pdf_stream, format='PDF', save_all=True, append_images=other_imgs)
    pdf_stream.seek(0)
    return pdf_stream