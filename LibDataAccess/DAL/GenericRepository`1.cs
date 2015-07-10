using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Linq.Expressions;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Core.Objects.DataClasses;
using System.Data.Entity.Infrastructure;
using System.Data.Entity.Core.Objects;
using System.Data.Entity.Infrastructure.Interception;
using EntityFramework.BulkInsert.Extensions;
using System.Data.Entity.Core.Metadata.Edm;

namespace LibDataAccess.DAL
{
    /// <summary>
    /// 泛型Repository類別, 提供CRUD的擴充方法
    /// </summary>
    /// <typeparam name="TEntity">Entity型別</typeparam>
    public class GenericRepository<TEntity>
        where TEntity : class
    {
        private static object _lock = new object();

        /// <summary>
        /// EntityName與EntitySetName的對應Dict
        /// </summary>
        private Dictionary<string, string> _entitySetNameDict = new Dictionary<string, string>();

        /// <summary>
        /// DbContext物件
        /// </summary>
        protected DbContext context;

        /// <summary>
        /// Entity的DbSet物件
        /// </summary>
        protected DbSet<TEntity> dbSet;

        /// <summary>
        /// 由DbContext取得內含的ObjectContext物件
        /// </summary>
        protected ObjectContext ObjContext
        {
            get
            {
                return ((IObjectContextAdapter)context).ObjectContext;
            }
        }

        /// <summary>
        /// 建構子, 需傳入特定資料庫的DbContext物件
        /// </summary>
        /// <param name="context"></param>
        public GenericRepository(DbContext context)
        {
            this.context = context;
            this.dbSet = context.Set<TEntity>();
        }

        /// <summary>
        /// 以sql script取得Entity, 需撈取該Entity的所有欄位, 且無法透過join來做到Include
        /// </summary>
        /// <param name="sql">要執行的sql, 可使用參數, 從@p0開始或使用SqlParameter指定參數名稱</param>
        /// <param name="parameters">參數值, 可使用SqlParameter</param>
        /// <returns>Entity的查詢結果集合</returns>
        public virtual IEnumerable<TEntity> GetBySql(string sql, params object[] parameters)
        {
            return dbSet.SqlQuery(sql, parameters);
        }

        /// <summary>
        /// 取得Entity的所有資料
        /// </summary>
        /// <param name="includes">要預先讀取的Property</param>
        /// <returns>Entity的查詢結果集合</returns>
        public virtual IQueryable<TEntity> GetAll(IEnumerable<Expression<Func<TEntity, object>>> includes = null)
        {
            IQueryable<TEntity> query = dbSet;

            if (includes != null)
            {
                query = includes.Aggregate(query, (current, include) => current.Include(include));
            }

            return query;
        }

        /// <summary>
        /// 依據條件搜尋Entity
        /// </summary>
        /// <param name="filter">搜尋條件</param>
        /// <param name="orderBy">排序條件</param>
        /// <param name="page">現在頁數</param>
        /// <param name="pageSize">每頁筆數</param>
        /// <param name="includeProperties">要預先讀取的Property</param>
        /// <returns>Entity的查詢結果集合</returns>
        public virtual IQueryable<TEntity> Get(
            Expression<Func<TEntity, bool>> filter = null,
            Func<IQueryable<TEntity>, IOrderedQueryable<TEntity>> orderBy = null,
            IEnumerable<Expression<Func<TEntity, object>>> includes = null,
            int page = -1,
            int pageSize = -1)
        {
            IQueryable<TEntity> query = dbSet;

            if (filter != null)
            {
                query = query.Where(filter);
            }

            if (includes != null)
            {
                query = includes.Aggregate(query, (current, include) => current.Include(include));
            }

            if (orderBy != null)
            {
                query = orderBy(query);
            }

            if (page > 0 && pageSize > 0)
            {
                int skipCount = (page - 1) * pageSize;
                query = query.Skip(skipCount);
            }

            if (pageSize > 0)
            {
                query = query.Take(pageSize);
            }

            return query;
        }

        /// <summary>
        /// 取得指定Entity於DbContext中的Entry資訊
        /// </summary>
        /// <param name="item">Entity</param>
        /// <returns>DbContext中的Entry資訊</returns>
        public virtual DbEntityEntry<TEntity> GetEntry(TEntity item)
        {
            return context.Entry<TEntity>(item);
        }

        /// <summary>
        /// 先從追蹤的物件搜尋, 若無再搜尋資料庫
        /// 故Add後尚未SaveChange前還是可以取的到
        /// </summary>
        /// <param name="keyValues">Entity的PK值, 需依照順序</param>
        /// <returns>Entity</returns>
        public virtual TEntity Find(params object[] keyValues)
        {
            return dbSet.Find(keyValues);
        }

        /// <summary>
        /// 先從追蹤的物件搜尋, 若無再搜尋資料庫
        /// 故Add後尚未SaveChange前還是可以取的到
        /// </summary>
        /// <param name="item">擁有PK值的Entity</param>
        /// <returns>Entity</returns>
        public virtual TEntity Find(TEntity item)
        {
            object[] keyValue = GetEntityKeyValue(item);
            TEntity entity = Find(keyValue);
            return entity;
        }

        /// <summary>
        /// 以具有EntityKey的Entity直接從資料庫取得完整Entity
        /// </summary>
        /// <param name="item">具有EntityKey值的Entity</param>
        /// <returns>完整Entity</returns>
        public virtual TEntity Load(TEntity item)
        {
            var keyValue = ObjContext.CreateEntityKey(typeof(TEntity).Name, item);

            object o = null;
            ObjContext.TryGetObjectByKey(keyValue, out o);

            return o as TEntity;
        }

        /// <summary>
        /// 更新DbSet中的Entity, 若Entity不存在則新增至DbSet中
        /// </summary>
        /// <param name="item">Entity</param>
        public virtual void AddOrUpdate(TEntity item)
        {
            DbEntityEntry<TEntity> entry = context.Entry<TEntity>(item);

            if (entry.State == EntityState.Detached)
            {
                object[] keyValue = GetEntityKeyValue(item);
                TEntity entity = Find(keyValue);

                if (entity != null)
                {
                    context.Entry(entity).CurrentValues.SetValues(item);
                }
                else
                {
                    Add(item);
                }
            }
        }

        /// <summary>
        /// 批次更新DbSet中的Entity, 若Entity不存在則新增至DbSet中
        /// </summary>
        /// <param name="items">Entity的集合</param>
        public virtual void AddOrUpdateMany(IEnumerable<TEntity> items)
        {
            foreach (TEntity item in items)
            {
                AddOrUpdate(item);
            }
        }

        /// <summary>
        /// 新增Entity至DbSet中
        /// </summary>
        /// <param name="item">Entity</param>
        public virtual void Add(TEntity item)
        {
            dbSet.Add(item);
        }

        /// <summary>
        /// 新增多筆Entity至DbSet中
        /// </summary>
        /// <param name="items">Entity的集合</param>
        public virtual void AddMany(IEnumerable<TEntity> items)
        {
            dbSet.AddRange(items);
        }

        /// <summary>
        /// 批次新增, 會直接寫入資料庫, 不須再呼叫UnitOfWork.SaveChanges()
        /// </summary>
        /// <param name="items">Entity的集合</param>
        public virtual void BulkInsert(IEnumerable<TEntity> items)
        {
            context.BulkInsert(items);
        }

        /// <summary>
        /// 更新DbSet中的Entity
        /// </summary>
        /// <param name="item">要更新的Entity</param>
        public virtual void Update(TEntity item)
        {
            DbEntityEntry<TEntity> entry = context.Entry<TEntity>(item);

            if (entry.State == EntityState.Detached)
            {
                object[] keyValue = GetEntityKeyValue(item);
                TEntity entity = Find(keyValue);

                if (entity != null)
                {
                    context.Entry(entity).CurrentValues.SetValues(item);
                }
            }
        }

        /// <summary>
        /// 批次更新DbSet中的Entity
        /// </summary>
        /// <param name="items">要更新的Entity集合</param>
        public virtual void UpdateMany(IEnumerable<TEntity> items)
        {
            foreach (TEntity item in items)
            {
                Update(item);
            }
        }

        /// <summary>
        /// 移除DbSet中的Entity
        /// </summary>
        /// <param name="item">要移除的Entity</param>
        public virtual void Remove(TEntity item)
        {
            DbEntityEntry<TEntity> entry = context.Entry<TEntity>(item);

            if (entry.State == EntityState.Detached)
            {
                object[] keyValue = GetEntityKeyValue(item);
                TEntity entity = Find(keyValue);

                if (entity != null)
                {
                    dbSet.Remove(entity);
                }
            }
            else
            {
                dbSet.Remove(item);
            }
        }

        /// <summary>
        /// 批次移除DbSet中的Entity
        /// </summary>
        /// <param name="items">要移除的Entity集合</param>
        public virtual void RemoveMany(IEnumerable<TEntity> items)
        {
            foreach (TEntity item in items)
            {
                Remove(item);
            }
        }

        /// <summary>
        /// 取得Entity的PK值
        /// </summary>
        /// <param name="item">Entity</param>
        /// <returns>PK值的集合</returns>
        protected object[] GetEntityKeyValue(TEntity item)
        {
            string entitySetName = GetEntitySetName(typeof(TEntity).Name);

            var key = ObjContext.CreateEntityKey(entitySetName, item);

            object[] keyValue = key.EntityKeyValues.Select(p => p.Value).ToArray();

            return keyValue;
        }

        /// <summary>
        /// 以EntityName取得EntitySet Name
        /// </summary>
        /// <param name="entityName"></param>
        /// <returns></returns>
        private string GetEntitySetName(string entityName)
        {
            string entitySetName = "";
            if (_entitySetNameDict.TryGetValue(entityName, out entitySetName) == false)
            {
                lock (_lock)
                {
                    var container = ObjContext.MetadataWorkspace.GetEntityContainer(ObjContext.DefaultContainerName, DataSpace.CSpace);
                    entitySetName = (from meta in container.BaseEntitySets
                                     where meta.ElementType.Name == entityName
                                     select meta.Name).First();

                    if (_entitySetNameDict.ContainsKey(entityName) == false)
                    {
                        _entitySetNameDict.Add(entityName, entitySetName);
                    }
                }
            }

            return entitySetName;
        }
    }
}
