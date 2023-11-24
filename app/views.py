from django.contrib.auth.decorators import login_required
from django.shortcuts import render
from django.http import HttpRequest, HttpResponse
from app.defToExcel import gerarExcel
from datetime import date


dataInicio = 0
dataFim = 0
minCodColaborador = 0
maxCodColaborador = 0
Cliente = 0
maxCliente = 0
Atividade = 0
maxAtividade = 0
CentroCusto = 0
maxCentroCusto = 0


def main(request: HttpRequest) -> HttpResponse:
    global dataInicio, dataFim, minCodColaborador, maxCodColaborador, Cliente, maxCliente, Atividade, maxAtividade, CentroCusto, maxCentroCusto
    if request.method == "GET":
        dataInicio = request.GET.get("DataInicio")
        dataFim = request.GET.get("DataFim")
        minCodColaborador = request.GET.get("CodColaborador")
        maxCodColaborador = request.GET.get("maxCodColaborador")
        Cliente = str(request.GET.get("Cliente"))
        maxCliente = str(request.GET.get("maxCliente"))
        Atividade = str(request.GET.get("Atividade"))
        maxAtividade = str(request.GET.get("maxAtividade"))
        CentroCusto = str(request.GET.get("CentroCusto"))
        maxCentroCusto = str(request.GET.get("maxCentroCusto"))
    return render(request, "global/app.html")


