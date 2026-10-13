# Lab 4 - XML and JSON

Lab 4 coursework: build an XML document and a JSON document, then display both in HTML pages
with the JavaScript `fetch` API.

Student: 乔润喆  247050

## Folder structure

```
Lab4-XML-JSON/
├── index.html              landing page linking the two demos
├── xml-lab/
│   ├── book.xml            the XML book catalogue (each book has an id attribute)
│   └── index.html          fetch + DOMParser + parseFromString(xmlText, 'text/xml')
├── json-lab/
│   ├── userProfiles.json   4 user profiles (address + preferences + subscription)
│   └── index.html          fetch + response.json()
├── reports/
│   ├── Lab4-247050-乔润喆.doc      the report to submit (file name required by the lab sheet)
│   ├── Lab4-247050-乔润喆.docx     the same report in DOCX format
│   ├── Lab4-247050-乔润喆.pdf      PDF preview of the report
│   └── screenshots/                screenshots used in the report
└── tools/
    └── build_report.py             script that rebuilds the report and trims the figures
```

### About the `.doc` file

The DOCX and the `.doc` file contain exactly the same report. Word opens the `.doc` file
normally (it is OOXML inside, so Word may ask once whether to open it). If you want a real
Word 97-2003 file, open `Lab4-247050-乔润喆.docx` in Word and choose
**文件 → 另存为 → Word 97-2003 文档 (*.doc)**. The PDF is only for quick viewing.

## Run it locally

`fetch()` cannot read local files from `file://`, so serve the folder over HTTP:

```
cd Lab4-XML-JSON
python -m http.server 8765
```

Then open <http://127.0.0.1:8765/xml-lab/index.html> and
<http://127.0.0.1:8765/json-lab/index.html>.

## Published on GitHub Pages

Live site: **<https://1396832296.github.io/lab4/>**

* XML experiment - <https://1396832296.github.io/lab4/xml-lab/>
* JSON experiment - <https://1396832296.github.io/lab4/json-lab/>

Repository: <https://github.com/1396832296/lab4> (Settings -> Pages -> Deploy from a branch,
`main` / `(root)`).

## What the code does (no libraries, only what the lab sheet asks for)

* **XML experiment** - `fetch("book.xml")` -> `response.text()` -> `new DOMParser()` ->
  `parser.parseFromString(xmlText, 'text/xml')` -> loop the `<book>` elements with
  `getElementsByTagName()` and show `id`, `title`, `author`, `price`. A `<parsererror>` check
  reports whether the XML is well-formed.
* **JSON experiment** - `fetch('userProfiles.json')` -> check `response.ok` -> `response.json()` ->
  loop the array and show `firstName`, `lastName`, `email`, `dateOfBirth`, `address` and
  `preferences` (theme, language, notifications, subscription).
