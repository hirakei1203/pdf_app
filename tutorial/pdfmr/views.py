from django.shortcuts import render

from django.views import generic
from django.contrib.auth.mixins import LoginRequiredMixin
from django.conf import settings
from .forms import UploadForm
from .utils import create_excel
from django.core.files.storage import default_storage
import shutil, os

# Create your views here.

def top(request):
    return render(request, 'pdfmr/top.html')
    
class UploadView(LoginRequiredMixin, generic.FormView):
    form_class = UploadForm
    template_name = 'pdfmr/upload_form.html'


    def form_valid(self, form):
        user_name = self.request.user.username  #ログオンユーザ名の取得
        user_dir = os.path.join(settings.MEDIA_ROOT, "excel", user_name)   #ユーザディレクトリパスの生成
        if not os.path.isdir(user_dir):  #ユーザディレクトリの作成
            os.makedirs(user_dir)
        temp_dir = form.save()  # upload一時フォルダの取得
        create_excel(temp_dir, user_name)
        shutil.rmtree(temp_dir)  #upload一時フォルダの削除
        _, file_list = default_storage.listdir(os.path.join(settings.MEDIA_ROOT, "excel", user_name))
        message = "正常終了しました。"
        context = {
            'file_list': file_list,
            'user_name':user_name,
            'message': message,
        }
        return render(self.request, 'pdfmr/complete.html', context)


    def form_invalid(self, form):
        return render(self.request, 'pdfmr/upload_form.html', {'form': form})
        
class ListView(LoginRequiredMixin, generic, TemplateView):
    
    template_name = 'pdfmr/list.html'
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        login_user_name = self.request.user.username
        file_list = default_storage.listdir(os.path.join(settings.MEDIA_ROOT, "excel", login_user_name))[1]
        context = {
            'file_list' : file_list,
            'login_user_name' : login_user_name,
        }
        return context