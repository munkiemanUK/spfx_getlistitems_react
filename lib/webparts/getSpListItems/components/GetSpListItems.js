var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from 'react';
import styles from './GetSpListItems.module.scss';
import { Web } from "@pnp/sp/presets/all";
import { DetailsList, SelectionMode, DetailsListLayoutMode, mergeStyles, Link, Image, ImageFit } from '@fluentui/react';
import "@pnp/sp/lists";
import "@pnp/sp/items";
var GetSpListItems = /** @class */ (function (_super) {
    __extends(GetSpListItems, _super);
    function GetSpListItems(props) {
        var _this = _super.call(this, props) || this;
        _this._onRenderItemColumn = function (item, index, column) {
            var src = item.ProductImage;
            switch (column.key) {
                case 'ProductImage':
                    return (React.createElement("a", { href: item.ProductImage, target: "_blank" },
                        React.createElement(Image, { src: src, width: 50, height: 50, imageFit: ImageFit.cover })));
                case 'ProductName':
                    return (React.createElement("span", { "data-selection-disabled": true, style: { whiteSpace: 'normal' } }, item.ProductName));
                case 'CustomerName':
                    return React.createElement(Link, { style: { whiteSpace: 'normal' }, href: "#" }, item.CustomerName);
                case 'CustomerEmail':
                    return React.createElement("span", { style: { whiteSpace: 'normal' } }, item.CustomerEmail);
                case 'OrderDate':
                    return (React.createElement("span", { "data-selection-disabled": true, className: mergeStyles({ height: '100%', display: 'block' }) }, item.OrderDate));
                case 'ProductDescription':
                    return (React.createElement("span", { "data-selection-disabled": true, style: { whiteSpace: 'normal' } }, item.ProductDescription));
                default:
                    return React.createElement("span", null, item.CustomerName);
            }
        };
        var columns = [
            {
                key: "ProductImage",
                name: "",
                fieldName: "ProductImage",
                minWidth: 70,
                maxWidth: 90,
                isResizable: true,
                data: "string",
                isPadded: true,
            },
            {
                key: "ProductName",
                name: "Product Name",
                fieldName: "ProductName",
                minWidth: 70,
                maxWidth: 90,
                isResizable: true,
                data: "string",
                isPadded: true,
            },
            {
                key: "CustomerName",
                name: "Customer Name",
                fieldName: "CustomerName",
                minWidth: 70,
                maxWidth: 90,
                isRowHeader: true,
                isResizable: true,
                data: "string",
                isPadded: true
            },
            {
                key: "CustomerEmail",
                name: "Customer Email",
                fieldName: "CustomerEmail",
                minWidth: 70,
                maxWidth: 90,
                isResizable: true,
                data: "number",
                isPadded: true
            },
            {
                key: "OrderDate",
                name: "Order Date",
                fieldName: "OrderDate",
                minWidth: 70,
                maxWidth: 90,
                isResizable: true,
                data: "string",
            },
            {
                key: "ProductDescription",
                name: "Product Description",
                fieldName: "ProductDescription",
                minWidth: 210,
                maxWidth: 350,
                isResizable: true,
                data: "string",
            }
        ];
        _this.state = {
            Items: [],
            columns: columns,
            isColumnReorderEnabled: true,
        };
        return _this;
    }
    GetSpListItems.prototype.componentWillMount = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.getData()];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    GetSpListItems.prototype.getData = function () {
        return __awaiter(this, void 0, void 0, function () {
            var data, web, items;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        data = [];
                        web = Web(this.props.webURL);
                        return [4 /*yield*/, web.lists.getByTitle("FluentUITest").items.select("*", "CustomerName/EMail", "CustomerName/Title").expand("CustomerName/ID").get()];
                    case 1:
                        items = _a.sent();
                        console.log(items);
                        return [4 /*yield*/, items.forEach(function (item) { return __awaiter(_this, void 0, void 0, function () {
                                return __generator(this, function (_a) {
                                    switch (_a.label) {
                                        case 0: return [4 /*yield*/, data.push({
                                                CustomerName: item.CustomerName.Title,
                                                CustomerEmail: item.CustomerName.EMail,
                                                ProductName: item.ProductName,
                                                OrderDate: FormatDate(item.OrderDate),
                                                ProductDescription: item.ProductDescription,
                                                ProductImage: window.location.origin + item.ProductImage.match('"serverRelativeUrl":(.*),"id"')[1].replace(/['"]+/g, '')
                                            })];
                                        case 1:
                                            _a.sent();
                                            return [2 /*return*/];
                                    }
                                });
                            }); })];
                    case 2:
                        _a.sent();
                        console.log(data);
                        return [4 /*yield*/, this.setState({ Items: data })];
                    case 3:
                        _a.sent();
                        console.log(this.state.Items);
                        return [2 /*return*/];
                }
            });
        });
    };
    GetSpListItems.prototype.render = function () {
        return (React.createElement("div", { className: styles.getSpListItems },
            React.createElement("h1", null, "Display SharePoint list data using spfx"),
            React.createElement(DetailsList, { items: this.state.Items, columns: this.state.columns, setKey: "set", layoutMode: DetailsListLayoutMode.justified, isHeaderVisible: true, onRenderItemColumn: this._onRenderItemColumn, selectionMode: SelectionMode.none }),
            React.createElement("h3", null, "Note : Image is clickable.")));
    };
    return GetSpListItems;
}(React.Component));
export default GetSpListItems;
export var FormatDate = function (date) {
    var date1 = new Date(date);
    var year = date1.getFullYear();
    var month = (1 + date1.getMonth()).toString();
    month = month.length > 1 ? month : '0' + month;
    var day = date1.getDate().toString();
    day = day.length > 1 ? day : '0' + day;
    return month + '/' + day + '/' + year;
};
//# sourceMappingURL=GetSpListItems.js.map