import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI, spfi, SPFx } from '@pnp/sp';

import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

export interface IExportData {
  Id: number;
  CountryName: string;
  Quantity: number;
  ExportValue: number;
  PercentageShare: number;
}

export class ExportService {

  private _sp: SPFI;
  private _listName: string = 'ExportDetails';

  constructor(context: WebPartContext) {

    this._sp = spfi()
      .using(
        SPFx(context)
      );
  }

  public async getExportDetails(): Promise<IExportData[]> {

    const items: any[] =
      await this._sp.web.lists
        .getByTitle(this._listName)
        .items
        .select(
          'Id',
          'Title',
          'Quantity',
          'ExportValue'
        )
        .top(5000)();

    const data: IExportData[] = items.map(
      item => ({
        Id: item.Id,
        CountryName: item.Title || '',
        Quantity: Number(item.Quantity) || 0,
        ExportValue: Number(item.ExportValue) || 0,
        PercentageShare: Number(item.PercentageShare) || 0
      })
    );

    const totalExportValue =
      data.reduce(
        (total, item) =>
          total + item.ExportValue,
        0
      );

    return data.map(item => ({
      ...item,
      PercentageShare:
        totalExportValue > 0
          ? (item.ExportValue / totalExportValue) * 100
          : 0
    }));
  }
}