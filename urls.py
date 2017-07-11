from django.conf.urls import url
from ibk.views import *
from . import views

urlpatterns = [
    #$를 붙이면 해당 주소 뒤로 다른 url이 추가되어도 각각 다 처리할 수 있고 $이 없다면 특정 주소 뒤로 어떠한 url이 오더라도 해당 url에서 지정한 view에서 모두 처리한다
    url(r'^$', upload_file, name='upload'),
    # ?p<parameter 이름>파라미터 값 형태로 url을 통해 파라미터를 넘길 수 있다
    url(r'^page/(?P<code>R(\d+)-(\d+))/', show_normal_report, name='page'),
    url(r'download/$', download, name='download'),
]