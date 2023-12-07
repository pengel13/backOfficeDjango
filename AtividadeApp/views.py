# from django.contrib.auth.decorators import login_required
from django.shortcuts import render, redirect
from django.http import HttpRequest, HttpResponse
from django.contrib import messages, auth
from django.contrib.auth.forms import AuthenticationForm
from django.contrib.auth.decorators import login_required
from datetime import datetime
from AtividadeApp.forms import FiltrosForm
from AtividadeApp.defToExcel import gerarExcel


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
            wb = gerarExcel(
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
            response = HttpResponse(content_type="application/ms-excel")

            today = datetime.now()
            responseContent = f'attachment; filename = AtividadeIntegracoes{today.hour}{today.minute}{today.hour} {today.day}-{today.month}-{today.year}.xlsx'
            
            response['Content-Disposition'] = responseContent

            wb.save(response) # type: ignore

            return response
    return render(request, "AtividadeApp/base.html", context={"form": filtros_form})

def login_view(request):
    form = AuthenticationForm(request)
    if request.method == "POST":
        form = AuthenticationForm(request, data=request.POST)
        if form.is_valid():
            messages.success(request, "USUÁRIO LOGADO")
            print("IS-VALIDD")
            user = form.get_user()
            auth.login(request, user)
            return redirect("main")
        messages.error(request, "Houve um erro no user ou na senha")
    context = {
        "form": form,
    }
    return render(request, "AtividadeApp/login.html", context)


def logout_view(request):
    auth.logout(request)
    messages.warning(request, 'USUÁRIO DESLOGADO')
    return redirect("login")


