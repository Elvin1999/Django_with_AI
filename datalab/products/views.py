import os

from django.conf import settings
from django.shortcuts import render

from . import utils
from .forms import UploadForm
from .models import Product
from django.db.models import Count,F,Sum,Avg,DecimalField,ExpressionWrapper

# Create your views here.

def dashboard(request):
    kpi=Product.objects.aggregate(
        products=Count('id'),
        total_qty=Sum("quantity"),
        avg_price=Avg("price"),
    )

    revenue_expr=ExpressionWrapper(F("price")*F("quantity"),
                                   output_field=DecimalField(max_digits=14,decimal_places=2))
    top_cats=(Product.objects
              .values("category")
              .annotate(revenue=Sum(revenue_expr),items=Count("id"))
              .order_by("-revenue")[:5])

    return render(request,"products/dashboard.html",{"kpi":kpi,"top_cats":top_cats})

def product_upload(request):
    ctx={"form":UploadForm()}
    if request.method=="POST":
        form=UploadForm(request.POST,request.FILES)
        if form.is_valid():
            up=request.FILES["file"]
            sheet=form.cleaned_data.get("sheet_name") or None

            updir=os.path.join(settings.MEDIA_ROOT,"uploads")
            os.makedirs(updir,exist_ok=True)
            fpath=os.path.join(updir,up.name)
            with open(fpath,"wb+") as dest:
                for ch in up.chunks():
                    dest.write(ch)

            df=utils.read_any(fpath,sheet)
            df=utils.normalize_for_product(df)

            rows=df.to_dict("records")
            if len(rows)==1:
                r=rows[0]

                Product.objects.update_or_create(
                    sku=r["sku"],
                    defaults=dict(
                        name=r["name"],
                        price=r["price"],
                        quantity=int(r["quantity"]),
                        category=r.get("category") or "",
                        tx_date=r["tx_date"],
                    )
                )

            elif len(rows)>1:
                for r in rows:
                    Product.objects.update_or_create(
                        sku=r["sku"],
                        defaults=dict(
                            name=r["name"],
                            price=r["price"],
                            quantity=int(r["quantity"]),
                            category=r.get("category") or "",
                            tx_date=r["tx_date"],
                        )
                    )

            ctx["msg"]=f"Uploaded : {len(rows)} rows"

    return render(request,"products/upload.html",ctx)