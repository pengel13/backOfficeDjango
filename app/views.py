from django.shortcuts import render
from django.http import HttpRequest, HttpResponse
from . import forms

def main(request: HttpRequest) -> HttpResponse:
    dataInicio= 0
    dataFim = 0
    minCodColaborador = 0
    maxCodColaborador = 0
    if request.method == 'GET':
        dataInicio = request.GET.get('DataInicio')
        dataFim = request.GET.get('DataFim')
        minCodColaborador = request.GET.get('CodColaborador')
        maxCodColaborador = request.GET.get('maxCodColaborador')
    print(dataInicio, dataFim, minCodColaborador, maxCodColaborador)
    return render(request, 'global/app.html')