import os
import io
import logging
from django.shortcuts import render
from django.http import HttpResponse
from django.conf import settings
from .utils import generate_pptx

logger = logging.getLogger(__name__)

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

    try:
        # Create an in-memory buffer to save the presentation without using the disk
        buffer = io.BytesIO()

        # Generate PPT and save it directly into the memory buffer
        generate_pptx(request.POST, request.FILES, {}, template_path, buffer)

        # Reset buffer position to the beginning before reading
        buffer.seek(0)

        response = HttpResponse(
            buffer,
            content_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        response["Content-Disposition"] = 'attachment; filename="single_page_presentation.pptx"'
        return response

    except Exception as e:
        logger.error("Error generating PPT", exc_info=True)
        return HttpResponse(f"An error occurred: {str(e)}", status=500)