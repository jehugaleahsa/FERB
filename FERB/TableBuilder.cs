﻿using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace FERB
{
    internal class TableBuilder<TModel>
        : ITableBuilder<TModel>,
        IOrderedTableBuilder<TModel>,
        INestedTableBuilder<TModel>,
        IWorksheetContent
    {
        private readonly OrderedDictionary headers;
        private readonly List<OrderByDefinition> orderBys;
        private readonly List<Func<TModel, int, int, IWorksheetContent>> nestedBuilders;
        private IWorksheetContentBuilder predecessor;
        private int startingRow;
        private int startingCell;
        private string title;
        private Action<ICellStyle> titleStyleApplier;
        private bool isHeaderHidden;
        private Action<ICellStyle> headerStyleApplier;
        private Action<ICellStyle> styleApplier;
        private IEnumerable<TModel> models;

        public TableBuilder()
        {
            this.headers = new OrderedDictionary();
            this.orderBys = new List<OrderByDefinition>();
            this.nestedBuilders = new List<Func<TModel, int, int, IWorksheetContent>>();
            this.startingRow = 1;
        }

        IWorksheetContentBuilder IWorksheetContentBuilder.Predecessor
        {
            get { return predecessor; }
        }

        public ITableBuilder<TModel> StartingAfter(IWorksheetContentBuilder builder)
        {
            this.predecessor = builder;
            return this;
        }

        IOrderedTableBuilder<TModel> IOrderedTableBuilder<TModel>.StartingAfter(IWorksheetContentBuilder builder)
        {
            StartingAfter(builder);
            return this;
        }

        public ITableBuilder<TModel> StartingAt(int rowNumber, int cellNumber)
        {
            if (rowNumber < 1)
            {
                throw new ArgumentOutOfRangeException("rowNumber", "The starting row number cannot be less than one.");
            }
            if (cellNumber < 0)
            {
                throw new ArgumentOutOfRangeException("cellNumber", "The starting cell number cannot be negative.");
            }
            this.startingRow = rowNumber;
            this.startingCell = cellNumber;
            return this;
        }

        IOrderedTableBuilder<TModel> IOrderedTableBuilder<TModel>.StartingAt(int rowNumber, int cellNumber)
        {
            StartingAt(rowNumber, cellNumber);
            return this;
        }

        public ITableBuilder<TModel> WithTitle(string text, Action<ICellStyle> applier = null)
        {
            this.title = text;
            this.titleStyleApplier = applier;
            return this;
        }

        IOrderedTableBuilder<TModel> IOrderedTableBuilder<TModel>.WithTitle(string text, Action<ICellStyle> applier)
        {
            WithTitle(text, applier);
            return this;
        }

        INestedTableBuilder<TModel> INestedTableBuilder<TModel>.WithTitle(string text, Action<ICellStyle> applier)
        {
            WithTitle(text, applier);
            return this;
        }

        public ITableBuilder<TModel> HideHeader(bool hide = true)
        {
            this.isHeaderHidden = hide;
            return this;
        }

        IOrderedTableBuilder<TModel> IOrderedTableBuilder<TModel>.HideHeader(bool hide)
        {
            HideHeader(hide);
            return this;
        }

        INestedTableBuilder<TModel> INestedTableBuilder<TModel>.HideHeader(bool hide)
        {
            HideHeader(hide);
            return this;
        }

        public ITableBuilder<TModel> WithHeaderStyleApplied(Action<ICellStyle> applier)
        {
            this.headerStyleApplier = applier;
            return this;
        }

        IOrderedTableBuilder<TModel> IOrderedTableBuilder<TModel>.WithHeaderStyleApplied(Action<ICellStyle> applier)
        {
            WithHeaderStyleApplied(applier);
            return this;
        }

        INestedTableBuilder<TModel> INestedTableBuilder<TModel>.WithHeaderStyleApplied(Action<ICellStyle> applier)
        {
            WithHeaderStyleApplied(applier);
            return this;
        }

        public ITableBuilder<TModel> WithStyleApplied(Action<ICellStyle> applier)
        {
            this.styleApplier = applier;
            return this;
        }

        IOrderedTableBuilder<TModel> IOrderedTableBuilder<TModel>.WithStyleApplied(Action<ICellStyle> applier)
        {
            WithStyleApplied(applier);
            return this;
        }

        INestedTableBuilder<TModel> INestedTableBuilder<TModel>.WithStyleApplied(Action<ICellStyle> applier)
        {
            WithStyleApplied(applier);
            return this;
        }

        public ITableBuilder<TModel> WithColumn(string headerName, Func<TModel, object> propertyAccessor, Action<IColumnConfiguration<TModel>> configurator = null)
        {
            if (String.IsNullOrWhiteSpace(headerName))
            {
                throw new ArgumentException("The header name cannot be blank.", "headerName");
            }
            if (propertyAccessor == null)
            {
                throw new ArgumentNullException("propertyAccessor");
            }
            ColumnConfiguration<TModel> configuration = new ColumnConfiguration<TModel>();
            if (configurator != null)
            {
                configurator(configuration);
            }
            headers[headerName] = new ColumnDefinition<TModel>()
            {
                HeaderName = headerName,
                Accessor = propertyAccessor,
                Configuration = configuration
            };
            return this;
        }

        IOrderedTableBuilder<TModel> IOrderedTableBuilder<TModel>.WithColumn(string headerName, Func<TModel, object> propertyAccessor, Action<IColumnConfiguration<TModel>> configurator)
        {
            WithColumn(headerName, propertyAccessor, configurator);
            return this;
        }

        INestedTableBuilder<TModel> INestedTableBuilder<TModel>.WithColumn(string headerName, Func<TModel, object> propertyAccessor, Action<IColumnConfiguration<TModel>> configurator)
        {
            WithColumn(headerName, propertyAccessor, configurator);
            return this;
        }

        public IOrderedTableBuilder<TModel> OrderedBy(string headerName, bool isDescending = false)
        {
            if (String.IsNullOrWhiteSpace(headerName))
            {
                return this;
            }
            OrderByDefinition definition = new OrderByDefinition() { HeaderName = headerName, IsDescending = isDescending };
            orderBys.Add(definition);
            return this;
        }

        IOrderedTableBuilder<TModel> IOrderedTableBuilder<TModel>.ThenBy(string headerName, bool isDescending)
        {
            if (String.IsNullOrWhiteSpace(headerName))
            {
                return this;
            }
            OrderByDefinition definition = new OrderByDefinition() { HeaderName = headerName, IsDescending = isDescending };
            orderBys.Add(definition);
            return this;
        }

        public ITableBuilder<TModel> WithData(IEnumerable<TModel> models)
        {
            this.models = models;
            return this;
        }

        IOrderedTableBuilder<TModel> IOrderedTableBuilder<TModel>.WithData(IEnumerable<TModel> models)
        {
            WithData(models);
            return this;
        }

        public ITableBuilder<TModel> WithNestedTable<TItem>(Func<TModel, IEnumerable<TItem>> propertyAccessor, Action<INestedTableBuilder<TItem>> configurator)
        {
            if (propertyAccessor == null)
            {
                throw new ArgumentNullException("propertyAccessor");
            }
            if (configurator == null)
            {
                throw new ArgumentNullException("configurator");
            }
            Func<TModel, int, int, IWorksheetContent> action = (model, startingRow, startingColumn) =>
            {
                IEnumerable<TItem> items = propertyAccessor(model);
                TableBuilder<TItem> builder = new TableBuilder<TItem>();
                builder.StartingAt(1, startingColumn + 1);
                builder.WithData(items);
                configurator(builder);
                return builder;
            };
            nestedBuilders.Add(action);
            return this;
        }

        IOrderedTableBuilder<TModel> IOrderedTableBuilder<TModel>.WithNestedTable<TItem>(Func<TModel, IEnumerable<TItem>> propertyAccessor, Action<INestedTableBuilder<TItem>> configurator)
        {
            WithNestedTable(propertyAccessor, configurator);
            return this;
        }

        int IWorksheetContent.Save(ExcelWorksheet worksheet, int rowOffset)
        {
            int currentRow = rowOffset + startingRow;
            if (this.headers.Count == 0)
            {
                // There are no columns, so we do not display anything
                return currentRow;
            }
            
            var columnDefinitions = headers.Keys.Cast<string>().Select(n => (ColumnDefinition<TModel>)headers[n]).ToArray();

            // Add title to table
            if (!String.IsNullOrWhiteSpace(title))
            {
                currentRow = addTitle(worksheet, currentRow, columnDefinitions);
            }

            // Apply any column configurations
            if (!isHeaderHidden)
            {
                TableHeaderRow<TModel> header = new TableHeaderRow<TModel>(startingCell, columnDefinitions);
                header.StyleApplier = headerStyleApplier;
                currentRow = header.Save(worksheet, currentRow);
                ++currentRow;
            }

            var records = sortModels();
            foreach (TModel model in records)
            {
                TableBodyRow<TModel> row = new TableBodyRow<TModel>(model, startingCell, columnDefinitions);
                row.StyleApplier = styleApplier;
                currentRow = row.Save(worksheet, currentRow);

                foreach (var nestedBuilder in nestedBuilders)
                {
                    IWorksheetContent content = nestedBuilder(model, currentRow, startingCell);
                    currentRow = content.Save(worksheet, currentRow);
                }

                ++currentRow;
            }
            --currentRow;
            return currentRow;
        }

        private int addTitle(ExcelWorksheet worksheet, int currentRow, IEnumerable<ColumnDefinition<TModel>> definitions)
        {
            int visibleColumnCount = definitions.Where(d => d.Configuration.IsVisible).Count();
            int lastCell = startingCell + visibleColumnCount;  // Subtract since the end cell is inclusive
            ExcelRange range = worksheet.Cells[currentRow, startingCell + 1, currentRow, lastCell];
            range.Merge = true;
            range.Value = title;
            if (titleStyleApplier != null)
            {
                titleStyleApplier(new CellStyle(range.Style));
            }
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ++currentRow;
            return currentRow;
        }

        private IEnumerable<TModel> sortModels()
        {
            if (models == null)
            {
                return Enumerable.Empty<TModel>();
            }
            if (!models.Any() || !orderBys.Any())
            {
                return models;
            }
            var records = models.OrderBy(t => 0);
            foreach (OrderByDefinition orderBy in orderBys)
            {
                var accessor = getOrderByAccessor(orderBy);
                if (accessor != null)
                {
                    if (orderBy.IsDescending)
                    {
                        records = records.ThenByDescending(accessor);
                    }
                    else
                    {
                        records = records.ThenBy(accessor);
                    }
                }
            }
            return records;
        }

        private Func<TModel, object> getOrderByAccessor(OrderByDefinition orderBy)
        {
            if (!headers.Contains(orderBy.HeaderName))
            {
                return null;
            }
            var definition = (ColumnDefinition<TModel>)headers[orderBy.HeaderName];
            return definition.Accessor;
        }

        private class OrderByDefinition
        {
            public string HeaderName { get; set; }

            public bool IsDescending { get; set; }
        }
    }
}
