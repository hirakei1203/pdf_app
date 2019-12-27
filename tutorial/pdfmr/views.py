from django.shortcuts import render

from django.views import generic
from django.contrib.auth.mixins import LoginRequiredMixin
from django.conf import settings
from .forms import UploadForm
from django.core.files.storage import default_storage
import shutil, os

# Create your views here.

def top(request):
    return render(request, 'pdfmr/top.html')
    
class UploadView(LoginRequiredMixin, generic.FormView)
    from_class = UploadForm
    template_name = 'pdfmr/upload_form.html'
    
    def form_valid(self, form)
        user_name = self.request.user.username
        user_dir = os.path.join(settings.MEDIA_ROOT, "excel", user_name)
        if not os.path.isdir(user_dir):
            os.makedirs(user_dir)
        temp_dir = form.save()
        # 変換実装予定
        shutil.rmtree(temp_dir)
        _, file_list = default_storage.listdir(os.path.join(settings.MEDIA_ROOT, "excel", user_name))
        message = "正常終了しました"
        context = {
            'file_list': file_list,
            'user_name': user_name,
            'message' : message,
        }
        return render(self.request, 'pdfmr/complete.html', context)
        
    def form_invalid(self, form):
        return render(self.request, 'pdfmr/uplaod_form.html', {'form': form})
        