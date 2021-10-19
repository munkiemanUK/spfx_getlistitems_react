import * as React from 'react';
import { IGetSpListItemsProps } from './IGetSpListItemsProps';
import { IColumn } from '@fluentui/react';
import "@pnp/sp/lists";
import "@pnp/sp/items";
export interface IGetSpListItemsStates {
    Items: IDocument[];
    columns: any;
    isColumnReorderEnabled: boolean;
}
export interface IDocument {
    CustomerName: string;
    CustomerEmail: string;
    ProductName: string;
    OrderDate: any;
    ProductDescription: any;
    ProductImage: string;
}
export default class GetSpListItems extends React.Component<IGetSpListItemsProps, IGetSpListItemsStates> {
    constructor(props: any);
    componentWillMount(): Promise<void>;
    getData(): Promise<void>;
    _onRenderItemColumn: (item: IDocument, index: number, column: IColumn) => JSX.Element | string;
    render(): React.ReactElement<IGetSpListItemsProps>;
}
export declare const FormatDate: (date: any) => string;
//# sourceMappingURL=GetSpListItems.d.ts.map