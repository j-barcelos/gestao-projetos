#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Busca de sabores responsiva — patch idempotente para o projeto Oggi.

Substitui a listagem crua de sabores por:
  - campo de busca com filtro instantâneo (ignora acento, caixa e ordem de palavras)
  - renderização limitada + botão "mostrar mais"
  - estado vazio, contador de resultados e botão de limpar
  - layout responsivo (mobile-first, sem zoom involuntário no iOS)

Nada é alterado em models, migrations ou API: o filtro é local, sobre o
cache já carregado por /api/sabores/.

Uso (dentro de meusite/):
    python aplicar_busca_sabores.py
    python aplicar_busca_sabores.py --check
    python aplicar_busca_sabores.py --reverter

Seguro para rodar repetidamente: blocos delimitados por marcadores são
substituídos no lugar, nunca duplicados.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

INICIO_JS = "/* == BUSCA_SABORES:START == */"
FIM_JS = "/* == BUSCA_SABORES:END == */"
INICIO_CSS = "/* == BUSCA_SABORES:START == */"
FIM_CSS = "/* == BUSCA_SABORES:END == */"
INICIO_HTML = "<!-- == BUSCA_SABORES:START == -->"
FIM_HTML = "<!-- == BUSCA_SABORES:END == -->"

REL_JS = Path("aluguel/static/aluguel/js/main.js")
REL_CSS = Path("aluguel/static/aluguel/css/style.css")
REL_HTML = Path("aluguel/templates/aluguel/index.html")

BLOCO_HTML = """            <!-- == BUSCA_SABORES:START == -->
            <div class="busca-sabores" id="busca-sabores-wrap" hidden>
                <div class="busca-sabores-campo">
                    <span class="busca-sabores-icone" aria-hidden="true">&#128269;</span>
                    <input
                        type="search"
                        id="busca-sabores"
                        class="busca-sabores-input"
                        placeholder="Buscar sabor..."
                        autocomplete="off"
                        spellcheck="false"
                        inputmode="search"
                        aria-label="Buscar sabor"
                        aria-controls="lista-sorvetes"
                    >
                    <button
                        type="button"
                        id="busca-sabores-limpar"
                        class="busca-sabores-limpar"
                        aria-label="Limpar busca"
                        hidden
                    >&times;</button>
                </div>
                <p class="busca-sabores-info" id="busca-sabores-info" role="status" aria-live="polite"></p>
            </div>
            <!-- == BUSCA_SABORES:END == -->
"""

BLOCO_CSS = """/* == BUSCA_SABORES:START == */
.busca-sabores {
    display: flex;
    flex-direction: column;
    gap: 8px;
    margin: 0 0 20px;
}

.busca-sabores[hidden] {
    display: none;
}

.busca-sabores-campo {
    display: flex;
    align-items: center;
    gap: 10px;
    background: #fff;
    border: 1.5px solid rgba(159, 48, 129, 0.18);
    border-radius: 14px;
    padding: 6px 8px 6px 14px;
    box-shadow: 0 6px 18px rgba(159, 48, 129, 0.06);
    transition: border-color .2s ease, box-shadow .2s ease;
}

.busca-sabores-campo:focus-within {
    border-color: var(--oggi-rosa);
    box-shadow: 0 0 0 4px rgba(230, 40, 102, 0.12);
}

.busca-sabores-icone {
    font-size: 16px;
    line-height: 1;
    opacity: .55;
}

.busca-sabores-input {
    flex: 1;
    min-width: 0;
    border: none;
    outline: none;
    background: transparent;
    font: inherit;
    font-size: 16px;
    color: var(--oggi-texto);
    padding: 10px 0;
}

.busca-sabores-input::-webkit-search-cancel-button {
    display: none;
}

.busca-sabores-limpar {
    flex-shrink: 0;
    width: 34px;
    height: 34px;
    border: none;
    border-radius: 50%;
    background: rgba(159, 48, 129, 0.1);
    color: var(--oggi-roxo);
    font-size: 20px;
    line-height: 1;
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: center;
    transition: background-color .2s ease, color .2s ease;
}

.busca-sabores-limpar:hover {
    background: var(--oggi-rosa);
    color: #fff;
}

.busca-sabores-limpar[hidden] {
    display: none;
}

.busca-sabores-info {
    margin: 0;
    min-height: 18px;
    font-size: 13px;
    font-weight: 700;
    color: var(--oggi-texto-suave);
}

.busca-sabores-vazio {
    grid-column: 1 / -1;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    gap: 6px;
    padding: 40px 20px;
    text-align: center;
    border: 2px dashed rgba(159, 48, 129, 0.2);
    border-radius: 16px;
    color: var(--oggi-texto-suave);
}

.busca-sabores-vazio strong {
    font-size: 17px;
    color: var(--oggi-roxo);
}

.busca-sabores-mais-wrap {
    grid-column: 1 / -1;
    display: flex;
    justify-content: center;
    padding: 6px 0 2px;
}

.busca-sabores-mais {
    padding: 12px 22px;
    border: 1.5px solid rgba(159, 48, 129, 0.25);
    border-radius: 999px;
    background: #fff;
    color: var(--oggi-roxo);
    font: inherit;
    font-size: 14px;
    font-weight: 800;
    cursor: pointer;
    transition: background-color .2s ease, color .2s ease, border-color .2s ease, transform .2s ease;
}

.busca-sabores-mais:hover {
    background: var(--oggi-rosa);
    border-color: var(--oggi-rosa);
    color: #fff;
    transform: translateY(-1px);
}

.card-sem-imagem {
    background: linear-gradient(135deg, rgba(159, 48, 129, 0.08), rgba(26, 157, 193, 0.08));
}

@media (max-width: 560px) {
    .busca-sabores-campo {
        padding: 4px 6px 4px 12px;
    }

    .busca-sabores-vazio {
        padding: 30px 16px;
    }

    .busca-sabores-mais {
        width: 100%;
    }
}
/* == BUSCA_SABORES:END == */
"""

BLOCO_JS = """/* == BUSCA_SABORES:START == */
(function () {
  'use strict';

  var LIMITE_PAGINA = 12;
  var MINIMO_PARA_BUSCA = 6;
  var DEBOUNCE_MS = 120;

  var termoAtual = '';
  var quantidadeVisivel = LIMITE_PAGINA;
  var timer = null;

  function normalizar(valor) {
    return String(valor == null ? '' : valor)
      .normalize('NFD')
      .replace(/[\\u0300-\\u036f]/g, '')
      .toLowerCase()
      .replace(/\\s+/g, ' ')
      .trim();
  }

  function escapeHtml(valor) {
    return String(valor == null ? '' : valor)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function listaSabores() {
    try {
      return Array.isArray(saboresCache) ? saboresCache : [];
    } catch (erro) {
      return [];
    }
  }

  function filtrar(lista, termo) {
    var alvo = normalizar(termo);
    if (!alvo) return lista;

    var palavras = alvo.split(' ').filter(Boolean);

    return lista.filter(function (sabor) {
      var nome = normalizar(sabor && sabor.nome);
      return palavras.every(function (palavra) {
        return nome.indexOf(palavra) !== -1;
      });
    });
  }

  function cardHtml(sabor) {
    var id = Number(sabor.id);
    var qtd = (carrinho[id] && carrinho[id].qtd) || 0;
    var preco = parseFloat(sabor.preco);
    var precoTxt = (isNaN(preco) ? 0 : preco).toFixed(2);
    var nome = escapeHtml(sabor.nome);

    var imagem = sabor.imagem_url
      ? '<img src="' + escapeHtml(sabor.imagem_url) + '" alt="' + nome + '" loading="lazy" decoding="async">'
      : '<img src="" alt="" class="card-sem-imagem" loading="lazy" decoding="async">';

    return '' +
      '<div class="card">' +
        imagem +
        '<h3>' + nome + '</h3>' +
        '<p>R$ ' + precoTxt + '</p>' +
        '<div class="qty-control">' +
          '<button type="button" class="btn-qty btn-menos" onclick="atualizarQtd(' + id + ', -1)" aria-label="Diminuir ' + nome + '">-</button>' +
          '<input type="number" id="qty-' + id + '" class="input-qty" value="' + qtd + '" min="0" max="999" inputmode="numeric" aria-label="Quantidade de ' + nome + '" oninput="validarLimite(this)" onchange="inputManual(' + id + ', this.value)">' +
          '<button type="button" class="btn-qty btn-mais" onclick="atualizarQtd(' + id + ', 1)" aria-label="Aumentar ' + nome + '">+</button>' +
        '</div>' +
      '</div>';
  }

  function atualizarInterface(todos, encontrados, exibidos) {
    var wrap = document.getElementById('busca-sabores-wrap');
    var info = document.getElementById('busca-sabores-info');
    var limpar = document.getElementById('busca-sabores-limpar');
    var buscando = normalizar(termoAtual).length > 0;

    if (wrap) wrap.hidden = todos.length <= MINIMO_PARA_BUSCA;
    if (limpar) limpar.hidden = !buscando;

    if (!info) return;

    if (!buscando) {
      info.textContent = todos.length
        ? 'Exibindo ' + exibidos.length + ' de ' + todos.length + ' sabores'
        : 'Nenhum sabor disponível no momento';
      return;
    }

    if (!encontrados.length) {
      info.textContent = 'Nenhum sabor corresponde à busca';
      return;
    }

    info.textContent = encontrados.length + ' sabor' +
      (encontrados.length > 1 ? 'es encontrados' : ' encontrado');
  }

  function renderizar() {
    var container = document.getElementById('lista-sorvetes');
    if (!container) return;

    var todos = listaSabores();
    var encontrados = filtrar(todos, termoAtual);
    var exibidos = encontrados.slice(0, quantidadeVisivel);

    var partes = exibidos.map(cardHtml);

    if (!exibidos.length) {
      partes.push(
        '<div class="busca-sabores-vazio">' +
          '<strong>Nenhum sabor encontrado</strong>' +
          '<span>Tente outro nome ou limpe a busca.</span>' +
        '</div>'
      );
    } else if (encontrados.length > exibidos.length) {
      var restantes = encontrados.length - exibidos.length;
      partes.push(
        '<div class="busca-sabores-mais-wrap">' +
          '<button type="button" class="busca-sabores-mais" id="busca-sabores-mais">' +
            'Mostrar mais ' + restantes + ' sabor' + (restantes > 1 ? 'es' : '') +
          '</button>' +
        '</div>'
      );
    }

    container.innerHTML = partes.join('');
    atualizarInterface(todos, encontrados, exibidos);
  }

  function limparBusca() {
    var input = document.getElementById('busca-sabores');
    termoAtual = '';
    quantidadeVisivel = LIMITE_PAGINA;
    if (input) {
      input.value = '';
      input.focus();
    }
    renderizar();
  }

  function iniciar() {
    var input = document.getElementById('busca-sabores');
    var limpar = document.getElementById('busca-sabores-limpar');
    var container = document.getElementById('lista-sorvetes');
    if (!input) return;

    input.addEventListener('input', function () {
      clearTimeout(timer);
      timer = setTimeout(function () {
        termoAtual = input.value;
        quantidadeVisivel = LIMITE_PAGINA;
        renderizar();
      }, DEBOUNCE_MS);
    });

    input.addEventListener('keydown', function (evento) {
      if (evento.key === 'Escape') limparBusca();
    });

    if (limpar) limpar.addEventListener('click', limparBusca);

    if (container) {
      container.addEventListener('click', function (evento) {
        var alvo = evento.target.closest('#busca-sabores-mais');
        if (!alvo) return;
        quantidadeVisivel += LIMITE_PAGINA;
        renderizar();
      });
    }
  }

  window.renderizarCards = renderizar;
  window.limparBuscaSabores = limparBusca;

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', iniciar);
  } else {
    iniciar();
  }
})();
/* == BUSCA_SABORES:END == */
"""

ANCORA_HTML = re.compile(r'^[ \t]*<div id="lista-sorvetes"', re.MULTILINE)


class FalhaDeAncora(RuntimeError):
    pass


def backup(caminho: Path) -> None:
    destino = caminho.with_suffix(caminho.suffix + ".bak")
    if not destino.exists():
        destino.write_bytes(caminho.read_bytes())


def padrao_bloco(inicio: str, fim: str) -> re.Pattern:
    return re.compile(re.escape(inicio) + r".*?" + re.escape(fim) + r"\n?", re.DOTALL)


def aplicar(texto: str, inicio: str, fim: str, bloco: str, ancora) -> tuple[str, str]:
    existente = padrao_bloco(inicio, fim)
    if existente.search(texto):
        return existente.sub(lambda _m: bloco, texto, count=1), "atualizado"

    if ancora is not None:
        achado = ancora.search(texto)
        if not achado:
            raise FalhaDeAncora("âncora não encontrada")
        return texto[: achado.start()] + bloco + "\n" + texto[achado.start():], "inserido"

    return texto.rstrip("\n") + "\n\n" + bloco, "anexado"


def reverter(texto: str, inicio: str, fim: str) -> tuple[str, str]:
    padrao = padrao_bloco(inicio, fim)
    if not padrao.search(texto):
        return texto, "ausente"
    return padrao.sub("", texto, count=1), "removido"


def processar(caminho: Path, inicio: str, fim: str, bloco: str, ancora, args) -> str:
    if not caminho.is_file():
        raise FileNotFoundError(f"arquivo não encontrado: {caminho}")

    original = caminho.read_text(encoding="utf-8")

    if args.reverter:
        novo, estado = reverter(original, inicio, fim)
    else:
        try:
            novo, estado = aplicar(original, inicio, fim, bloco, ancora)
        except FalhaDeAncora:
            raise FalhaDeAncora(f"não foi possível localizar o ponto de inserção em {caminho}")

    if estado != "ausente" and novo.count(inicio) > 1:
        raise RuntimeError(f"marcador duplicado em {caminho}")

    if not args.check and novo != original:
        backup(caminho)
        caminho.write_text(novo, encoding="utf-8")

    return estado


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Aplica a busca de sabores responsiva de forma idempotente."
    )
    parser.add_argument("--root", default=str(Path(__file__).resolve().parent),
                        help="raiz do projeto Django (padrão: pasta do script)")
    parser.add_argument("--check", action="store_true", help="simula sem gravar nada")
    parser.add_argument("--reverter", action="store_true", help="remove os blocos aplicados")
    args = parser.parse_args()

    raiz = Path(args.root).resolve()
    if not (raiz / "manage.py").is_file():
        print(f"[erro] manage.py não encontrado em {raiz}", file=sys.stderr)
        return 2

    alvos = [
        ("html", raiz / REL_HTML, INICIO_HTML, FIM_HTML, BLOCO_HTML, ANCORA_HTML),
        ("css", raiz / REL_CSS, INICIO_CSS, FIM_CSS, BLOCO_CSS, None),
        ("js", raiz / REL_JS, INICIO_JS, FIM_JS, BLOCO_JS, None),
    ]

    falhou = False
    for rotulo, caminho, inicio, fim, bloco, ancora in alvos:
        try:
            estado = processar(caminho, inicio, fim, bloco, ancora, args)
        except (FileNotFoundError, FalhaDeAncora, RuntimeError) as erro:
            falhou = True
            print(f"[erro] {rotulo:<4} {caminho.name}: {erro}", file=sys.stderr)
            continue

        prefixo = "simularia" if args.check else estado
        print(f"[{prefixo}] {rotulo:<4} {caminho.relative_to(raiz)}")

    if falhou:
        return 1

    if args.check:
        print("\nNada foi gravado (--check).")
    elif args.reverter:
        print("\nBlocos removidos. Backups .bak preservados.")
    else:
        print("\nConcluído. Rode novamente quando quiser: o resultado não duplica.")
        print("Reinicie o runserver e limpe o cache do navegador (Ctrl+F5).")

    return 0


if __name__ == "__main__":
    sys.exit(main())