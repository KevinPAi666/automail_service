# myapp/views.py
from django.shortcuts import render
from django.http import HttpResponse
import os


def home(request):
    print('service trigger')
    exec(open("./automail.py", encoding='utf-8').read(), globals())
    return HttpResponse("報名/確認信寄出完成!")
