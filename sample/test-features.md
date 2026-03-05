---
ispras_templates:
  header_ru: 'Тестовые возможности конвертера'
  header_en: 'Converter Feature Tests'

  authors:
    - name_ru: 'И.И. Иванов'
      name_en: 'I.I. Ivanov'
      orcid: '0000-0000-0000-0000'
      email: '<ivanov@ispras.ru>'
      organizations: [ispras]
      details_ru: >-
        Иван Иванович ИВАНОВ – тестовый автор.
      details_en: >-
        Ivan Ivanovich IVANOV – test author.

  organizations:
    - id: ispras
      name_ru: 'Институт системного программирования им. В.П. Иванникова РАН, Россия, 109004, г. Москва, ул.
        А. Солженицына, д. 25.'
      name_en: 'Ivannikov Institute for System Programming of the Russian Academy of Sciences, 25,
        Alexander Solzhenitsyn st., Moscow, 109004, Russia.'

  bibliography: test-features.bib

  abstract_ru: >-
    Данный документ предназначен для тестирования возможностей конвертера
    proceedings-md: листингов, формул, списков.
  abstract_en: >-
    This document is intended for testing proceedings-md converter features:
    listings, formulas, lists.

  keywords_ru: 'листинги; формулы; списки'
  keywords_en: 'listings; formulas; lists'

  acknowledgements_ru: '@none'
  acknowledgements_en: '@none'
---

## 1. Листинги

Фрагмент кода программного продукта оформляется в виде листинга. Подписи
должны быть на двух языках и начинаться с текста вида \"Листинг 1\"
(\"Listing 1\"). Для написания программного кода используется шрифт
«Courier new» прямым начертанием (не курсив), обычный (нежирный). Размер
шрифта 9 пт. Ссылки на листинг в тексте статьи должны иметь вид
\"листинг <span class=ref>lst:lst1</span>\".

```rust
fn write(f: &File, data: &[u8]) -> io::Result<()> {
   f.write_at(0, data)?;
   f.ensure_durable(0..data.len())
}
```

<div class="listing-caption">Листинг <span class=ref>lst:lst1</span>. Пример листинга</div>
<div class="listing-caption">Listing <span class=ref>lst:lst1</span>. Listing example</div>

## 2. Формулы

Все формулы набираются с помощью формульного редактора. Формулы
располагаются по центру. Если формулы нумеруются, то их номера
заключаются в круглые скобки и располагаются с правого края:

$$\begin{array}{r}
U_{1} = n_{1}n_{1} + \frac{n_{1}\left( n_{1} + 1 \right)}{2} - R_{1};\#(1)
\end{array}$$

## 3. Списки

Списки выравниваются «по ширине», выравнивание на 0 см, отступ текста:
0,6 см.

Все пункты маркированного списка имеют одинаковые маркеры, они
отображаются в виде маленьких чёрных кругов:

-   пример;
-   маркированного;
-   списка.

В нумерованном списке вначале идет число, затем закрывающая скобка:

1)  пример;
2)  нумерованного;
3)  списка.

Пример ссылки на литературу [@Dijkstra1976].

# Список литературы / References
