from django.urls import path
from . import views

urlpatterns = [
    path('make_excel/<str:sheet_name>/', views.create_excel, name='make_excel'),
    path('make_daily_summary_sheet/<str:sheet_name>/',views.create_daily_summary_sheet, name='make_daily_summary_sheet'),

]
