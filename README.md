# JSXlsxCore
xlsx core for javascript,can use in node or browser.

使用JavaScript编写的`微软Office Excel格式xlsx`数据内核，可以用于node或者现代浏览器。

如果要保存为xlsx文件，需要搭配“xlsx-saver”模块（[git](https://github.com/jifengg/JSXlsxSaver)/[npm](https://www.npmjs.com/package/xlsx-saver)）使用。

# use in node

shell:
```shell
npm i js-xlsx-core
```

js:

```javascript
const XlsxCore = require('js-xlsx-core');
const {Book,Sheet} = XlsxCore;
var book = new Book();
var sheet = book.CreateSheet("第一页");
//...
```

# use in browser

shell:
```shell
npm i js-xlsx-core
```

html:
```html
<script src="node_modules/js-xlsx-core/xlsxcore.js"></script>
```

js:

```javascript
//xlsxcore.js auto add XlsxCore to window
const {Book,Sheet} = window.XlsxCore;
var book = new Book();
var sheet = book.CreateSheet("第一页");
//...
```


# demo

更多使用方式参看[这里](https://github.com/jifengg/JSXlsxDemo)
