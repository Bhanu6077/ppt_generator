from django.urls import path
from . import views

urlpatterns = [
    # Notice how we use views.generate here to match the new code
    path('', views.generate_ppt, name='generate_ppt'), 
]