from django.urls import path
from . import views
urlpatterns = [
    path("",views.dashboard,name="dashboard"),
    # path("products/",views.product_list,name="product_list"),
    path("products/upload/",views.product_upload,name="product_upload"),
    # path("products/export/",views.product_export,name="product_export"),
    # path("stats/",views.stats_view,name="stats"),
]