
(() => {
    /**
     * 一个Excel数据
     */
    function Book() {
        /**
         * 所有的页数据
         * @type {{[x:string]:Sheet}}
         */
        this.Sheets = {};
        /**
         * 整个Excel文件中的默认单元格样式
         * @type {CellStyle}
         */
        this.DefaultCellStyle = Object.assign(new CellStyle(), { FontSize: 11, FontName: '微软雅黑' });

        //以下变量为程序所使用的唯一id值，不要在外部进行更改！

        this.__sheetIndex = 0;
        this.__cellStyleIndex = 0;
        this.__shareStringIndex = 0;
        this.__numberFormatIndex = 176;
        this.__imageIndex = 0;
        this.__hyperlinkIndex = 0;
        this.__cellFontIndex = 0;
        this.__cellFillIndex = 0;
    }

    /**
     * 创建一个数据页
     * @param {string} name 名称
     * @returns {Sheet}
     */
    Book.prototype.CreateSheet = function (name) {
        if (this.Sheets[name] != null) {
            throw '已存在同名的sheet。[' + name + ']';
        }
        this.__sheetIndex++;
        const sheet = new Sheet(name, 'customSheet' + this.__sheetIndex);
        sheet.Book = this;
        this.Sheets[name] = sheet;
        return sheet;
    }

    /**
     * 根据名称获取一个数据页
     * @param {string} name 名称
     * @returns {Sheet}
     */
    Book.prototype.Sheet = function (name) {
        return this.Sheets[name];
    }

    /**
     * 保存为一个node Buffer，该Buffer保存之后就是一个xlsx文件。
     */
    Book.prototype.SaveAsBuffer = async function () {
        const zipData = this.MakeXlsx(this);
        var content = await zipData.generateAsync({ type: 'nodebuffer', compression: "DEFLATE", compressionOptions: { level: 5 } });
        return content;
    }

    /**
     * 保存为一个browser bolb，该bolb下载到本地后就是一个xlsx文件。
     */
    Book.prototype.SaveAsBolb = async function () {
        const zipData = this.MakeXlsx(this);
        var content = await zipData.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 5 } });
        return content;
    }

    /**
     * 生成一个xlsx格式的zip文件。该方法需要加入“xlsxsaver.js”以自动实现
     */
    Book.prototype.MakeXlsx = function () {
        throw 'MakeXlsx Not init!';
    }

    /**
     * 创建一个共用的单元格样式，多处使用时可以减少文档体积。
     * @param {CellStyle} data
     */
    Book.prototype.CreateShareCellStyle = function (data) {
        var style = new CellStyle();
        style = Object.assign(style, data);
        style.__id = this.__cellStyleIndex++;
        return style;
    }

    /**
     * 使用CreateShareString创建在文档中多处使用的文本，可以减少文档的体积。如果是数字请不要用ShareString
     */
    Book.prototype.CreateShareString = function (txt) {
        return new ShareString(txt, this.__shareStringIndex++);
    }

    /**
     * 创建一个通用的数字格式化方式，多处使用时可以减少文档体积。
     */
    Book.prototype.CreateShareNumberFormat = function (code) {
        this.__numberFormatIndex++;
        var format = new NumberFormat(this.__numberFormatIndex, code);
        return format;
    }

    /**
     * 创建一个图片，多处使用时可以减少文档体积。
     * @param {Buffer|string} imgData 图片数据，支持Base64字符串或者Buffer
     * @param {ImageOption} opt 图片选项
     */
    Book.prototype.CreateImage = function (imgData, opt) {
        this.__imageIndex++;
        var image = new Image(this.__imageIndex, imgData, opt);
        return image;
    }

    /**
     * 创建一个超链接，这里的样式会覆盖之前设置的样式，多处使用时可以减少文档体积。
     * @param {string} url 链接
     * @param {CellStyle} style 单元格样式
     */
    Book.prototype.CreateHyperlink = function (url, style) {
        this.__hyperlinkIndex++;
        var link = new Hyperlink(this.__hyperlinkIndex, url, style);
        return link;
    }

    /**
     * 创建一个共用的单元格字体信息，多处使用时可以减少文档体积。
     * @param {CellFont} data
     */
    Book.prototype.CreateShareCellFont = function (data) {
        var font = new CellFont();
        font = Object.assign(font, data);
        font.__id = this.__cellFontIndex++;
        return font;
    }

    /**
     * 创建一个共用的单元格填充信息，多处使用时可以减少文档体积。
     * @param {CellFill} data
     */
    Book.prototype.CreateShareCellFill = function (data) {
        var fill = new CellFill();
        fill = Object.assign(fill, data);
        fill.__id = this.__cellFillIndex++;
        return fill;
    }

    /**
     * 一个Excel页
     * @param {string} name 名称
     * @param {string} id 唯一id
     */
    function Sheet(name, id) {
        this.Name = name;
        this.id = id;
        /**
         * 所有的单元格数据
         * @type {{[x:number]:{[x:number]:Cell}}}
         */
        this.Datas = {};
        /**
         * 默认列宽
         */
        this.DefaultWidth = 10;
        /**
         * 默认行高
         */
        this.DefaultHeight = 16;
        /**
         * 所属的Book
         * @type {Book}
         */
        this.Book = null;
        /**
         * 图片列表
         * @type {[{img:Image,col:Number,row:number,width:Number,height:Number}]}
         */
        this.ImageList = [];

        //以下变量为存储的数据，请不要在外部改变。

        this.colWidth = {};
        this.rowHeight = {};
        this.mergeCellDatas = [];
    }

    /**
     * 添加一个文本信息到指定单元格中。
     * @param {string|number|Date|ShareString} txt 内容，支持Number，string，ShareString
     * @param {number} row 行 
     * @param {number} col 列 
     * @param {CellStyle} style 文本选项，允许为null，为null则使用默认样式。
     */
    Sheet.prototype.AddText = function (txt, row, col, style) {

        var rowData = this.Datas[row];
        if (rowData == null) {
            rowData = {};
            this.Datas[row] = rowData;
        }
        if (typeof txt == 'string') {
            txt = this.Book.CreateShareString(txt);
        }
        var cell = new Cell(txt);
        cell.Style = style;
        rowData[col] = cell;
        return cell;
    }

    /**
     * 添加一个文本
     */
    Sheet.prototype.AddImage = function (img, row, col, width, height) {
        this.ImageList.push({
            img,
            col,
            row,
            width,
            height
        });
    }

    /**
     * 设置列宽度
     * @param {number} col 列下标
     * @param {number} width 宽度
     */
    Sheet.prototype.SetColWidth = function (col, width) {
        this.colWidth[col] = width;
    }

    /**
     * 设置行高度
     * @param {number} row 行号
     * @param {number} height 高度
     */
    Sheet.prototype.SetRowHeight = function (row, height) {
        this.rowHeight[row] = height;
        this.Datas[row] = this.Datas[row] || {};
    }

    /**
     * 合并单元格
     * @param {number} fromCol 要合并的起始单元格所在的列
     * @param {number} fromRow 要合并的起始单元格所在的行
     * @param {number} toCol 要合并的结束单元格所在的列
     * @param {number} toRow 要合并的结束单元格所在的行
     */
    Sheet.prototype.MergeCell = function (fromRow, fromCol, toRow, toCol) {
        this.mergeCellDatas.push({ fromCol, fromRow, toCol, toRow });
    }

    /**
     * 共享文本。如果多个单元格存在相同文本，使用共享文本可以减少文档体积。
     * @param {string} txt 文本内容
     * @param {number} id 全局唯一id
     */
    function ShareString(txt, id) {
        this.__id = id;
        this.txt = txt;
    }

    /**
     * 一个单元格
     * @param {string|number|Date|ShareString} text 内容
     */
    function Cell(text) {
        /**
         * 单元格的显示内容
         */
        this.Text = text;
        /**
         * @type {CellStyle}
         */
        this.Style = null;
        /**
         * @type {Hyperlink}
         */
        this.Hyperlink = null;
    }

    /**
     * 单元格样式。可以设置多个属性：粗体、斜体、下划线、字体、字号、文字颜色、单元格背景色、文字对齐方式、数字显示格式
     */
    function CellStyle() {
        /**
         * 全局唯一id，自行创建CellStyle时不要赋值。推荐使用Book.prototype.CreateShareCellStyle来创建样式。
         */
        this.__id = null;
        /**
         * 字体
         * @type {CellFont}
         */
        this.Font = null;
        /**
         * 填充信息
         * @type {CellFill}
         */
        this.Fill = null;
        /**
         * 文字对齐方式
         * @type {CellAlignment}
         */
        this.Alignment = null;
        /**
         * 数字显示格式
         * @type {NumberFormat}
         */
        this.Format = null;
    }

    /**
     * 单元格中文字的显示字体
     */
    function CellFont() {
        /**
         * 全局唯一id，自行创建CellFont时不要赋值。推荐使用Book.prototype.CreateShareCellFont来创建样式。
         */
        this.__id = null;
        /**
         * 是否显示为粗体
         * @type {boolean}
         */
        this.Bold = false;
        /**
         * 是否显示为下划线
         * @type {boolean}
         */
        this.Underline = false;
        /**
         * 是否显示为斜体
         * @type {boolean}
         */
        this.Italic = false;
        /**
         * 字体名字
         * @type {string}
         */
        this.FontName = null;
        /**
         * 字号
         * @type {number}
         */
        this.FontSize = null;
        /**
         * 文字颜色
         * @type {string|number} FFFFFF字符串格式的颜色码，或者是颜色主题
         */
        this.Color = null;
    }

    /**
     * 单元格的填充信息
     */
    function CellFill() {
        /**
         * 全局唯一id，自行创建CellFill时不要赋值。推荐使用Book.prototype.CreateShareCellFill来创建样式。
         */
        this.__id = null;
        /**
         * 背景颜色
         * @type {string} FFFFFF字符串格式的颜色码
         */
        this.BGColor = null;
    }

    /**
     * 单元格中文字的对齐方式
     */
    function CellAlignment() {
        /**
         * 是否支持换行,默认为false
         * @type {boolean}
         */
        this.WrapText = false;
        /**
         * 水平对齐方式。可选值为HorizontalAlignment的所有值
         */
        this.Horizontal = null;
        /**
         * 垂直对齐方式。可选值为VerticalAlignment的所有值
         */
        this.Vertical = null;
    }

    /**
     * 数字格式方式
     * @param {number} id 全局唯一id
     * @param {string} code 自定义格式码
     */
    function NumberFormat(id, code) {
        /**
         * 全局唯一id，自行创建NumberFormat时不要赋值。建议采用Book.prototype.CreateShareNumberFormat创建。
         */
        this.__id = id;
        /**
         * @type {string} 数字的格式化方式，支持自定义。更多自定义码请执行搜索。
         */
        this.Code = code;
    }

    /**
     * 图片信息
     * @param {number} id 全局唯一id
     * @param {string|Buffer} data 图片数据，支持base64或者Buffer格式
     * @param {ImageOption} opt 选项
     */
    function Image(id, data, opt) {
        /**
         * 全局唯一id，自行创建时不要赋值。建议使用Book.prototype.CreateImage创建
         */
        this.__id = id;
        /**
         * 图片数据，支持base64或者Buffer格式
         * @type {string|Buffer}
         */
        this.Data = data;
        /**
         * 图片选项
         * @type {ImageOption}
         */
        this.Option = opt;
    }

    /**
     * 图片选项
     * @param {string} type 数据类型
     * @param {string} format 图片格式
     */
    function ImageOption(type, format) {
        /**
         * 数据类型，如：'base64'、'buffer'
         * @type {string} 
         */
        this.Type = type;
        /**
         * 图片格式，如'png'、'jpg'
         * @type {string} 
         */
        this.Format = format;
    }

    /**
     * 超链接信息
     * @param {number} id 全局唯一id
     * @param {string} url 超链接地址
     * @param {CellStyle} style 超链接文本的样式
     */
    function Hyperlink(id, url, style) {
        /**
         * 全局唯一id，自行创建时不要赋值。建议使用Book.prototype.CreateHyperlink创建。
         */
        this.__id = id;
        /**
         * 超链接地址
         * @type {string}
         */
        this.Link = url;
        /**
         * 超链接文本的样式
         * @type {CellStyle}
         */
        this.Style = style;
    }

    /**
     * 水平对齐方式
     */
    const HorizontalAlignment = {
        /**
         * 左对齐
         */
        Left: 'left',
        /**
         * 右对齐
         */
        Right: 'right',
        /**
         * 水平居中对齐
         */
        Center: 'center'
    }

    /**
     * 垂直对齐方式
     */
    const VerticalAlignment = {
        /**
         * 顶部对齐
         */
        Top: 'top',
        /**
         * 底部对齐
         */
        Bottom: 'bottom',
        /**
         * 垂直居中对齐
         */
        Center: 'center'
    }

    const XlsxCore = {
        Book,
        Sheet,
        Cell,
        ShareString,
        CellStyle,
        CellFont,
        CellFill,
        CellAlignment,
        NumberFormat,
        Image,
        ImageOption,
        HorizontalAlignment,
        VerticalAlignment
    }

    if (typeof global == 'undefined') {
        window.XlsxCore = XlsxCore;
    } else {
        module.exports = XlsxCore;
    }
})()