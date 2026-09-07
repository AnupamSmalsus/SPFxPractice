import * as React from 'react';
import { useEffect, useMemo, useState } from 'react';
import {
  ComposableMap,
  Geographies,
  Geography
} from 'react-simple-maps';

import styles from './ExportDetailsPortal.module.scss';

import {
  IExportDetailsPortalProps
} from './IExportDetailsPortalProps';

import {
  ExportService,
  IExportData
} from '../services/ExportService';

const geoUrl =
  'https://cdn.jsdelivr.net/npm/world-atlas@2/countries-110m.json';

interface ITooltip {
  visible: boolean;
  x: number;
  y: number;
  countryName: string;
  quantity: number;
  exportValue: number;
  percentageShare: number;
}

const countryNameMap: {
  [key: string]: string;
} = {
  'United States of America': 'United States',
  'United States': 'United States',
  'Russian Federation': 'Russia',
  'Russia': 'Russia',
  'Czechia': 'Czech Republic',
  'Czech Republic': 'Czech Republic',
  'Korea': 'South Korea',
  'Republic of Korea': 'South Korea',
  'South Korea': 'South Korea',
  'Iran': 'Iran',
  'Iran (Islamic Republic of)': 'Iran',
  'Viet Nam': 'Vietnam',
  'Vietnam': 'Vietnam',
  'Lao People’s Democratic Republic': 'Laos',
  'Lao PDR': 'Laos',
  'Laos': 'Laos',
  'Türkiye': 'Turkey',
  'Turkey': 'Turkey',
  'United Republic of Tanzania': 'Tanzania',
  'Tanzania': 'Tanzania',
  'Bolivia (Plurinational State of)': 'Bolivia',
  'Bolivia': 'Bolivia',
  'Venezuela (Bolivarian Republic of)': 'Venezuela',
  'Venezuela': 'Venezuela',
  'Brunei Darussalam': 'Brunei',
  'Brunei': 'Brunei',
  'Myanmar': 'Myanmar',
  'Burma': 'Myanmar',
  'Syrian Arab Republic': 'Syria',
  'Syria': 'Syria',
  'Côte d’Ivoire': 'Ivory Coast',
  "Cote d'Ivoire": 'Ivory Coast',
  'Ivory Coast': 'Ivory Coast',
  'Eswatini': 'Eswatini',
  'Swaziland': 'Eswatini',
  'North Macedonia': 'North Macedonia',
  'Macedonia': 'North Macedonia'
};

const normalizeCountryName = (
  countryName: string
): string => {

  const cleanedName =
    countryName
      .trim()
      .replace(/\s+/g, ' ');

  return (
    countryNameMap[cleanedName] ||
    cleanedName
  );
};

/*
 * ExportValue is assumed to be stored
 * as actual USD.
 *
 * Example:
 *
 * 50,000,000  = 50 USD Million
 * 100,000,000 = 100 USD Million
 * 500,000,000 = 500 USD Million
 * 1,000,000,000 = 1,000 USD Million
 */
const getExportValueInMillion = (
  exportValue: number
): number => {

  return exportValue / 1000000;

};

/*
 * Determines the color of the country
 * based on Export Value in USD Million.
 */
const getCountryColor = (
  exportValue: number
): string => {

  const valueInMillion =
    getExportValueInMillion(
      exportValue
    );

  if (valueInMillion > 1000) {
    return '#08306B';
  }

  if (valueInMillion >= 500) {
    return '#2171B5';
  }

  if (valueInMillion >= 100) {
    return '#4292C6';
  }

  if (valueInMillion >= 50) {
    return '#9ECAE1';
  }

  return '#DEEBF7';
};

const formatNumber = (
  value: number
): string => {

  return new Intl.NumberFormat(
    'en-IN'
  ).format(value);

};

const formatCurrency = (
  value: number
): string => {

  return new Intl.NumberFormat(
    'en-US',
    {
      maximumFractionDigits: 2,
      minimumFractionDigits: 2
    }
  ).format(value);

};

const ExportDetailsPortal:
  React.FC<IExportDetailsPortalProps> =
  props => {

    const [
      exportData,
      setExportData
    ] = useState<IExportData[]>([]);

    const [
      loading,
      setLoading
    ] = useState<boolean>(true);

    const [
      error,
      setError
    ] = useState<string>('');

    const [
      tooltip,
      setTooltip
    ] = useState<ITooltip>({
      visible: false,
      x: 0,
      y: 0,
      countryName: '',
      quantity: 0,
      exportValue: 0,
      percentageShare: 0
    });

    const exportService =
      useMemo(
        () =>
          new ExportService(
            props.Context
          ),
        [props.Context]
      );

    useEffect(() => {

      const loadExportData =
        async (): Promise<void> => {

          try {

            setLoading(true);
            setError('');

            const data =
              await exportService
                .getExportDetails();

            setExportData(data);

          } catch (err) {

            console.error(
              'Error loading export data:',
              err
            );

            setError(
              'Unable to load export details from SharePoint.'
            );

          } finally {

            setLoading(false);

          }

        };

      loadExportData();

    }, [exportService]);

    /*
     * Convert SharePoint data into a
     * country lookup object.
     */
    const exportLookup =
      useMemo(() => {

        const lookup: {
          [key: string]: IExportData;
        } = {};

        exportData.forEach(
          item => {

            const normalized =
              normalizeCountryName(
                item.CountryName
              );

            lookup[
              normalized.toLowerCase()
            ] = item;

          }
        );

        return lookup;

      }, [exportData]);

    const totalQuantity =
      useMemo(
        () =>
          exportData.reduce(
            (
              total,
              item
            ) =>
              total +
              item.Quantity,
            0
          ),
        [exportData]
      );

    const totalExportValue =
      useMemo(
        () =>
          exportData.reduce(
            (
              total,
              item
            ) =>
              total +
              item.ExportValue,
            0
          ),
        [exportData]
      );

    const handleMouseEnter = (
      event: React.MouseEvent,
      countryName: string
    ): void => {

      const normalizedCountry =
        normalizeCountryName(
          countryName
        );

      const data =
        exportLookup[
          normalizedCountry.toLowerCase()
        ];

      if (!data) {

        setTooltip(
          previous => ({
            ...previous,
            visible: false
          })
        );

        return;

      }

      setTooltip({
        visible: true,
        x: event.clientX + 15,
        y: event.clientY + 15,
        countryName:
          data.CountryName,
        quantity:
          data.Quantity,
        exportValue:
          data.ExportValue,
        percentageShare:
          data.PercentageShare
      });

    };

    const handleMouseMove = (
      event: React.MouseEvent
    ): void => {

      if (!tooltip.visible) {
        return;
      }

      setTooltip(
        previous => ({
          ...previous,
          x: event.clientX + 15,
          y: event.clientY + 15
        })
      );

    };

    const handleMouseLeave = (): void => {

      setTooltip(
        previous => ({
          ...previous,
          visible: false
        })
      );

    };

    if (loading) {

      return (
        <div
          className={
            styles.exportDetailsPortal
          }
        >

          <div
            className={
              styles.loadingContainer
            }
          >

            <div
              className={
                styles.spinner
              }
            />

            <span>
              Loading export details...
            </span>

          </div>

        </div>
      );

    }

    if (error) {

      return (
        <div
          className={
            styles.exportDetailsPortal
          }
        >

          <div
            className={
              styles.errorContainer
            }
          >

            {error}

          </div>

        </div>
      );

    }

    return (

      <div
        className={
          styles.exportDetailsPortal
        }
      >

        <div
          className={
            styles.header
          }
        >

          <div>

            <h2>
              Export Details
            </h2>

            <p>
              Export distribution by country
            </p>

          </div>

        </div>

        <div
          className={
            styles.summaryContainer
          }
        >

          <div
            className={
              styles.summaryCard
            }
          >

            <span>
              Countries
            </span>

            <strong>
              {exportData.length}
            </strong>

          </div>

          <div
            className={
              styles.summaryCard
            }
          >

            <span>
              Total Quantity
            </span>

            <strong>
              {formatNumber(
                totalQuantity
              )}
            </strong>

          </div>

          <div
            className={
              styles.summaryCard
            }
          >

            <span>
              Total Export Value
            </span>

            <strong>
              $
              {formatCurrency(
                totalExportValue
              )}
            </strong>

          </div>

        </div>

        <div
          className={
            styles.mapContainer
          }
        >

          <ComposableMap
            projection="geoEqualEarth"
            projectionConfig={{
              scale: 145
            }}
            className={
              styles.worldMap
            }
          >

            <Geographies
              geography={geoUrl}
            >

              {({
                geographies
              }) =>

                geographies.map(
                  geography => {

                    const countryName =
                      geography.properties.name;

                    const normalized =
                      normalizeCountryName(
                        countryName
                      );

                    const data =
                      exportLookup[
                        normalized.toLowerCase()
                      ];

                    const hasData =
                      !!data;

                    const countryColor =
                      hasData
                        ? getCountryColor(
                            data.ExportValue
                          )
                        : '#E5E5E5';

                    return (

                      <Geography
                        key={
                          geography.rsmKey ||
                          geography.id
                        }

                        geography={
                          geography
                        }

                        onMouseEnter={
                          event =>
                            handleMouseEnter(
                              event,
                              countryName
                            )
                        }

                        onMouseMove={
                          handleMouseMove
                        }

                        onMouseLeave={
                          handleMouseLeave
                        }

                        style={{
                          default: {
                            fill:
                              countryColor,
                            outline:
                              'none',
                            stroke:
                              '#FFFFFF',
                            strokeWidth:
                              0.5
                          },

                          hover: {
                            fill:
                              countryColor,
                            outline:
                              'none',
                            stroke:
                              hasData
                                ? '#333333'
                                : '#FFFFFF',
                            strokeWidth:
                              hasData
                                ? 1.2
                                : 0.5,
                            cursor:
                              hasData
                                ? 'pointer'
                                : 'default'
                          },

                          pressed: {
                            fill:
                              countryColor,
                            outline:
                              'none',
                            stroke:
                              '#333333',
                            strokeWidth:
                              1
                          }
                        }}
                      />

                    );

                  }
                )

              }

            </Geographies>

          </ComposableMap>

          <div
            className={
              styles.mapLegend
            }
          >

            <div
              className={
                styles.legendTitle
              }
            >
              Export Value (USD Million)
            </div>

            <div
              className={
                styles.legendItem
              }
            >

              <span
                className={
                  styles.legendColor
                }
                style={{
                  backgroundColor:
                    '#DEEBF7'
                }}
              />

              <span>
                0 – 50
              </span>

            </div>

            <div
              className={
                styles.legendItem
              }
            >

              <span
                className={
                  styles.legendColor
                }
                style={{
                  backgroundColor:
                    '#9ECAE1'
                }}
              />

              <span>
                50 – 100
              </span>

            </div>

            <div
              className={
                styles.legendItem
              }
            >

              <span
                className={
                  styles.legendColor
                }
                style={{
                  backgroundColor:
                    '#4292C6'
                }}
              />

              <span>
                100 – 500
              </span>

            </div>

            <div
              className={
                styles.legendItem
              }
            >

              <span
                className={
                  styles.legendColor
                }
                style={{
                  backgroundColor:
                    '#2171B5'
                }}
              />

              <span>
                500 – 1,000
              </span>

            </div>

            <div
              className={
                styles.legendItem
              }
            >

              <span
                className={
                  styles.legendColor
                }
                style={{
                  backgroundColor:
                    '#08306B'
                }}
              />

              <span>
                &gt; 1,000
              </span>

            </div>

            <div
              className={
                styles.legendItem
              }
            >

              <span
                className={
                  styles.noDataColor
                }
              />

              <span>
                No Data
              </span>

            </div>

          </div>

        </div>

        {tooltip.visible && (

          <div
            className={
              styles.tooltip
            }
            style={{
              left: tooltip.x,
              top: tooltip.y
            }}
          >

            <div
              className={
                styles.tooltipHeader
              }
            >

              {tooltip.countryName}

            </div>

            <div
              className={
                styles.tooltipRow
              }
            >

              <span>
                Quantity
              </span>

              <strong>
                {formatNumber(
                  tooltip.quantity
                )}
              </strong>

            </div>

            <div
              className={
                styles.tooltipRow
              }
            >

              <span>
                Export Value
              </span>

              <strong>
                $
                {formatCurrency(
                  tooltip.exportValue
                )}
              </strong>

            </div>

            <div
              className={
                styles.tooltipRow
              }
            >

              <span>
                USD Million
              </span>

              <strong>
                $
                {getExportValueInMillion(
                  tooltip.exportValue
                ).toFixed(2)}
              </strong>

            </div>

            <div
              className={
                styles.tooltipRow
              }
            >

              <span>
                Percentage Share
              </span>

              <strong>
                {tooltip.percentageShare.toFixed(
                  2
                )}
                %
              </strong>

            </div>

          </div>

        )}

        {!exportData.length && (

          <div
            className={
              styles.noData
            }
          >

            No export data is available.

          </div>

        )}

      </div>

    );

  };

export default ExportDetailsPortal;