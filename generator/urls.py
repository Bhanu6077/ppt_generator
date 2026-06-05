from django.urls import path
from . import views

urlpatterns = [
    # Notice how we use views.generate here to match the new code
    path('', views.generate_ppt, name='generate_ppt'), 
    path('history/', views.view_history, name='history'),
    path('edit/<int:id>/', views.edit_presentation, name='edit'),
    path('download/<int:id>/', views.download_presentation, name='download'),
]