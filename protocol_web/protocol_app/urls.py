from django.urls import path
from . import views

app_name = 'protocol_app'

urlpatterns = [
    path('', views.index, name='index'),
    path('upload/', views.upload_protocols, name='upload_protocols'),
    path('search/', views.search_protocols, name='search_protocols'),
    path('protocol/<int:pk>/', views.protocol_detail, name='protocol_detail'),
    path('protocol/<int:pk>/export/', views.export_single_protocol, name='export_single_protocol'),
    path('protocol/<int:pk>/delete/', views.delete_protocol, name='delete_protocol'),
    path('export/', views.export_protocols, name='export_protocols'),
    path('ajax/search/', views.ajax_search_protocols, name='ajax_search_protocols'),
    # path('admin/', admin.site.urls),
    path('export/', views.export_page, name='export_page'),
    path('export-to-excel/', views.export_to_excel, name='export_to_excel'),
]