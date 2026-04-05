import time

from desafio_ax import baixar_planilhas
import enviar_relatorio


def executar_processo():
    print("INICIANDO AUTOMAÇÃO COMPLETA...\n")

    try:
        # ─── ETAPA 1: DOWNLOAD ───────────
        print("Baixando planilhas...")
        baixar_planilhas()
        print("Download concluído!\n")

        time.sleep(2)

        # ─── ETAPA 2: RELATÓRIO ───────────
        print("Gerando relatório...")

        import importlib
        import relatorio  # IMPORTA SÓ AQUI (DEPOIS DO DOWNLOAD)
        importlib.reload(relatorio)

        print("Relatório gerado!\n")

        time.sleep(2)

        # ─── ETAPA 3: EMAIL ───────────
        print("Enviando email...")
        enviar_relatorio.main()

        print("PROCESSO FINALIZADO COM SUCESSO!")

    except Exception as e:
        print("ERRO NA AUTOMAÇÃO:")
        print(e)


if __name__ == "__main__":
    executar_processo()