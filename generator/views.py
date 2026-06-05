import os
import io
import tempfile
import logging
import json
from django.shortcuts import render, get_object_or_404
from django.http import HttpResponse
from django.conf import settings
from django.utils import timezone
from .utils import generate_pptx, convert_pptx_to_pdf
from .models import PresentationHistory
from django.shortcuts import render, get_object_or_404, redirect

logger = logging.getLogger(__name__)

def _get_template_path():
    return os.path.join(settings.BASE_DIR, "static", "ppt_templates", "master_template.pptx")

# GENEREATE VIEW
def generate_ppt(request):
    if request.method != "POST":
        return render(request, "generator/form.html")

    template_path = _get_template_path()
    history_id = request.POST.get("history_id")
    
    # 1. User ka naam secure karo (taaki space ki jagah special char error na de)
    raw_filename = request.POST.get("filename", "RoadAthena_Deck")
    safe_filename = "".join([c for c in raw_filename if c.isalnum() or c in (' ', '-', '_')]).strip()
    if not safe_filename:
        safe_filename = "RoadAthena_Deck"

    try:
        # 1. UPDATE ya CREATE logic
        if history_id:
            history_entry = PresentationHistory.objects.get(id=history_id)
            history_entry.name = safe_filename
            history_entry.payload = request.POST.dict()
            history_entry.created_at = timezone.now()
            history_entry.save()
        else:
            history_entry = PresentationHistory.objects.create(
                name=safe_filename,
                file_path=safe_filename,
                payload=request.POST.dict()
            )
            
        save_only = request.POST.get("save_only")
        if save_only == "true":
            # Agar user ne History wale button se save kiya hai, toh bina file download kare History pe bhej do
            return redirect('history')

        # 3. PPT / PDF GENERATION (Sirf tab chalega jab download par click kiya ho)
        export_format = request.POST.get("export_format", "pptx")
        # 3. PPT GENERATION
        if export_format == "pdf":
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_pptx_path = os.path.join(temp_dir, "temp_deck.pptx")
                with open(temp_pptx_path, "wb") as f:
                    generate_pptx(request.POST, request.FILES, {}, template_path, f)
                pdf_path = convert_pptx_to_pdf(temp_pptx_path, temp_dir)
                if pdf_path:
                    with open(pdf_path, 'rb') as pdf_file:
                        response = HttpResponse(pdf_file.read(), content_type='application/pdf')
                        response['Content-Disposition'] = f'attachment; filename="{safe_filename}.pdf"'
                        return response
        else:
            buffer = io.BytesIO()
            generate_pptx(request.POST, request.FILES, {}, template_path, buffer)
            buffer.seek(0)
            response = HttpResponse(buffer, content_type="application/vnd.openxmlformats-officedocument.presentationml.presentation")
            response["Content-Disposition"] = f'attachment; filename="{safe_filename}.pptx"'
            return response
            
    except Exception as e:
        logger.error("Error generating Presentation", exc_info=True)
        return HttpResponse(f"An error occurred: {str(e)}", status=500)

# HISTORY VIEW
def view_history(request):
    history = PresentationHistory.objects.all().order_by('-created_at')
    return render(request, "generator/history.html", {'history': history})

# EDIT VIEW
def edit_presentation(request, id):
    history_item = get_object_or_404(PresentationHistory, id=id)
    # Payload aur ID ke sath original naam bhi bhej rahe hain
    return render(request, "generator/form.html", {
        'payload': history_item.payload, 
        'payload_id': history_item.id,
        'original_name': history_item.name  # <-- YE NAYI LINE ADD KI HAI
    })


def download_presentation(request, id):
    history_item = get_object_or_404(PresentationHistory, id=id)
    template_path = _get_template_path()
    
    # Payload se data wapas lo
    data = history_item.payload
    
    buffer = io.BytesIO()
    # Generate PPT using stored payload
    generate_pptx(data, {}, {}, template_path, buffer)
    buffer.seek(0)
    
    response = HttpResponse(
        buffer,
        content_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
    response["Content-Disposition"] = f'attachment; filename="{history_item.name}.pptx"'
    return response