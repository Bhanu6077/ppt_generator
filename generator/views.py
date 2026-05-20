import os
import tempfile
from django.shortcuts import render
from django.http import HttpResponse
from django.conf import settings
from .utils import generate_pptx

# Template ka path nikalne ka safe tareeka
def _get_template_path():
    return os.path.join(
        settings.BASE_DIR, "static", "ppt_templates", "master_template.pptx"
    )

def generate_ppt(request):
    # Agar user sirf page open kar raha hai (GET request)
    if request.method != "POST":
        return render(request, "generator/form.html")

    # Agar user ne form submit kiya hai (POST request)
    template_path = _get_template_path()
    
    if not os.path.exists(template_path):
        return HttpResponse(f"Template file not found at: {template_path}. Please check your folder structure.", status=500)

    output_path = None
    try:
        # Ek temporary file banayenge jisme final PPT save hogi
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
            output_path = tmp.name

        # YAHAN CHANGE KIYA HAI: {} ki jagah request.FILES pass kiya hai
        generate_pptx(request.POST, request.FILES, {}, template_path, output_path)

        # File ko read karke user ko download ke liye bhej do
        with open(output_path, "rb") as f:
            data = f.read()

        response = HttpResponse(
            data,
            content_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        response["Content-Disposition"] = 'attachment; filename="single_page_presentation.pptx"'
        return response

    except Exception as e:
        return HttpResponse(f"An error occurred: {str(e)}", status=500)

    finally:
        # Server se temporary file delete kar do taaki kachra jama na ho
        if output_path and os.path.exists(output_path):
            os.unlink(output_path)