using System;

namespace FERB
{
    internal class ColumnDefinition<TModel>
    {
        public string HeaderName { get; set; }

        public Func<TModel, object> Accessor { get; set; }

        public ColumnConfiguration<TModel> Configuration { get; set; }
    }
}
