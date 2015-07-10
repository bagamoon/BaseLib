using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace LibWinCommon.Collection
{
    /// <summary>
    /// 可排序的List
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class SortedBindingList<T> : BindingList<T>
    {
        List<T> _list;
        ListSortDirection sortDirection;
        PropertyDescriptor sortProperty;
        Action<SortedBindingList<T>, List<T>> populateBaseList = (a, b) => a.ResetItems(b);
        static Dictionary<string, Func<List<T>, IEnumerable<T>>> cachedOrderByExpressions = new Dictionary<string, Func<List<T>, IEnumerable<T>>>();

        public SortedBindingList()
        {
            _list = new List<T>();
        }

        public SortedBindingList(IEnumerable<T> enumerable)
        {
            _list = enumerable.ToList();
            populateBaseList(this, _list);
        }

        public SortedBindingList(List<T> list)
        {
            _list = list;
            populateBaseList(this, _list);
        }

        protected override void ApplySortCore(PropertyDescriptor prop, ListSortDirection direction)
        {
            sortProperty = prop;
            sortDirection = direction;

            var orderByMethodName = sortDirection == ListSortDirection.Ascending ? "OrderBy" : "OrderByDescending";
            var cacheKey = typeof(T).GUID + prop.Name + orderByMethodName;

            if (!cachedOrderByExpressions.ContainsKey(cacheKey))
            {
                CreateOrderByMethod(prop, orderByMethodName, cacheKey);
            }

            ResetItems(cachedOrderByExpressions[cacheKey](_list).ToList());

            ResetBindings();
        }

        private void CreateOrderByMethod(PropertyDescriptor prop, string orderByMethodName, string cacheKey)
        {
            var sourceParameter = Expression.Parameter(typeof(List<T>), "source");
            var lambdaParameter = Expression.Parameter(typeof(T), "lambdaParameter");
            var accesedMember = typeof(T).GetProperty(prop.Name);
            var propertySelectorLambda = Expression.Lambda(Expression.MakeMemberAccess(lambdaParameter, accesedMember), lambdaParameter);
            var orderByMethod = typeof(Enumerable).GetMethods()
                                                  .Where(a => a.Name == orderByMethodName &&
                                                               a.GetParameters().Length == 2)
                                                  .Single()
                                                  .MakeGenericMethod(typeof(T), prop.PropertyType);

            var orderByExpression = Expression.Lambda<Func<List<T>, IEnumerable<T>>>(
                                        Expression.Call(orderByMethod,
                                                        new Expression[] { sourceParameter, 
                                                                           propertySelectorLambda }),
                                                        sourceParameter);

            cachedOrderByExpressions.Add(cacheKey, orderByExpression.Compile());
        }

        protected override void RemoveSortCore()
        {
            ResetItems(_list);
        }

        private void ResetItems(List<T> items)
        {
            base.ClearItems();

            for (int i = 0; i < items.Count; i++)
            {
                base.InsertItem(i, items[i]);
            }
        }

        protected override bool SupportsSortingCore
        {
            get
            {
                return true;
            }
        }

        protected override ListSortDirection SortDirectionCore
        {
            get
            {
                return sortDirection;
            }
        }

        protected override PropertyDescriptor SortPropertyCore
        {
            get
            {
                return sortProperty;
            }
        }

        protected override void OnListChanged(ListChangedEventArgs e)
        {

            _list = base.Items.ToList();
        }
    }
}
