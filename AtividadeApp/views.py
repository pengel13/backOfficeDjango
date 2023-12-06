# from django.contrib.auth.decorators import login_required
from django.shortcuts import render, redirect
from django.http import HttpRequest, HttpResponse
from AtividadeApp.defToExcel import gerarExcel
from django.contrib import messages, auth
from django.contrib.auth.forms import AuthenticationForm
from django.contrib.auth.decorators import login_required

from AtividadeApp.forms import FiltrosForm


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

@login_required
def main(request: HttpRequest) -> HttpResponse:
    filtros_form = FiltrosForm()
    if request.method == "POST":
        print("POST")
        filtros_form = FiltrosForm(request.POST)
        if filtros_form.is_valid():
            dados = filtros_form.cleaned_data
            dataInicio = dados.get("dataInicio")
            dataFim = dados.get("dataFim")
            minCodColaborador = dados.get("minCodColaborador")
            maxCodColaborador = dados.get("maxCodColaborador")
            minCliente = dados.get("minCliente")
            maxCliente = dados.get("maxCliente")
            minAtividade = dados.get("minAtividade")
            maxAtividade = dados.get("maxAtividade")
            minCentroCusto = dados.get("minCentroCusto")
            maxCentroCusto = dados.get("maxCentroCusto")
            stream = gerarExcel(
                dataInicio,
                dataFim,
                minCodColaborador,
                maxCodColaborador,
                minCliente,
                maxCliente,
                minAtividade,
                maxAtividade,
                minCentroCusto,
                maxCentroCusto,
            )
            # response = HttpResponse(content_type='application/vnd.ms-excel')
            # wb.save(response)
            # response = HttpResponse(save_workbook(wb, "AtividadeIntegracao.xlsx"), content_type='application/vnd.ms-excel')
            return stream

    return render(request, "AtividadeApp/base.html", context={"form": filtros_form})

def login_view(request):
    form = AuthenticationForm(request)
    print('n')
    if request.method == 'POST':
        print('POST')
        form = AuthenticationForm(request, data=request.POST)
        if form.is_valid():
            print('IS-VALIDD')
            user = form.get_user()
            auth.login(request, user)
            return redirect('main')
        
    context = {'form': form,}
    return render(
        request,
        'AtividadeApp/login.html',
        context
    )

def logout_view(request):
    auth.logout(request)
    return redirect('login')
