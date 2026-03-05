# Proceedings

Этот репозиторий содержит скрипт для автоматической конвертации `.md` в `.docx`,
соответствующий требованиям форматирования Трудов Института системного
программирования РАН.

## How to use:

```bash
git clone https://github.com/evdenis/proceedings-md
cd proceedings-md
npm install
npm run build
sudo apt-get install pandoc
```

Файл `sample.md` содержит стандартный шаблон статьи для Трудов ИСП РАН,
представленный в `.md`-формате. Скрипт `src/main.js` выполняет конвертацию.

```
cd sample
node ../src/main.js sample.md sample.docx
````

## Figures, Tables, and Listings

Use `@ref:prefix:label` for auto-numbered cross-references, and fenced divs (`:::`) for captions:

```markdown
![](image.png){width="4.7in" height="1in"}

::: img-caption
Рис. @ref:fig:fig1. Описание рисунка.
:::

::: img-caption
Fig. @ref:fig:fig1. Figure description.
:::
```

References are auto-numbered per prefix (`fig`, `tab`, `lst`). The same label always resolves to the same number — use it in both captions and body text:

```markdown
Как показано на рис. @ref:fig:fig1, ...
```

Supported caption classes: `img-caption`, `table-caption`, `listing-caption`.

## Notes

Скрипт несколько сырой. Ошибки могут быть нечитаемыми. Некоторые версии Microsoft Word
ругаются на то, что документ повреждён, но все равно открывают его. Перед отправкой
документа рекомендуется открыть документ в Word, перепроверить форматирование, и сохранить
заново.

Happy researching!
