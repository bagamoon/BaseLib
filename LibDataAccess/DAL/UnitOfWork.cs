using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Entity;
using System.Data;
using System.Data.Entity.Infrastructure;
using System.Data.Entity.Core.Objects;
using System.Data.Common;
using log4net;

namespace LibDataAccess.DAL
{
    /// <summary>
    /// UnitOfWork物件, 可使不同Repository共用同一DbContext
    /// </summary>
    public class UnitOfWork : IDisposable
    {
        /// <summary>
        /// transaction物件
        /// </summary>
        private DbTransaction _transaction;

        /// <summary>
        /// Dbcontext物件
        /// </summary>
        public DbContext context;

        /// <summary>
        /// Repository的Dictionary集合
        /// </summary>
        protected IDictionary<Type, object> repositoryDict;

        /// <summary>
        /// Log4net物件
        /// </summary>
        protected ILog logger;

        /// <summary>
        /// 由DbContext取得內含的ObjectContext物件
        /// </summary>
        public ObjectContext ObjContext
        {
            get
            {
                return ((IObjectContextAdapter)context).ObjectContext;
            }
        }

        /// <summary>
        /// 建構子, 需傳入特定資料庫的DbContext物件
        /// </summary>
        /// <param name="dbContext"></param>
        public UnitOfWork(DbContext dbContext)
        {
            context = dbContext;
            repositoryDict = new Dictionary<Type, object>();
            logger = LogManager.GetLogger(typeof(UnitOfWork).Name);
        }

        /// <summary>
        /// 以Entity取得Repository的實體, 會以Entity型別為Key存放於記憶體中
        /// </summary>
        /// <typeparam name="TEntity">Entity型別</typeparam>
        /// <returns>Repository</returns>
        public GenericRepository<TEntity> Repository<TEntity>() where TEntity : class
        {
            object repo = null;
            if (repositoryDict.TryGetValue(typeof(TEntity), out repo) == false)
            {
                repo = new GenericRepository<TEntity>(context);
                repositoryDict[typeof(TEntity)] = repo;
            }

            return repo as GenericRepository<TEntity>;
        }

        /// <summary>
        /// 直接執行sql script, 並回傳受影響筆數
        /// </summary>
        /// <param name="sql">要執行的sql, 可使用參數, 從@p0開始或使用SqlParameter指定參數名稱</param>
        /// <param name="parameters">參數值, 可使用SqlParameter</param>
        /// <returns>受影響筆數</returns>
        public virtual int ExecuteSqlCommand(string sql, params object[] parameters)
        {
            return context.Database.ExecuteSqlCommand(sql, parameters);
        }

        /// <summary>
        /// 直接執行sql script, 並回傳指定的型別, 屬性與select欄位需匹配
        /// </summary>
        /// <typeparam name="TEntity">指定型別</typeparam>
        /// <param name="sql">sql string</param>
        /// <param name="parameters">參數</param>
        /// <returns></returns>
        public IEnumerable<TEntity> SqlQuery<TEntity>(string sql, params object[] parameters)
        {
            return context.Database.SqlQuery<TEntity>(sql, parameters).AsEnumerable();
        }

        /// <summary>
        /// 將context追蹤的變更更新至資料庫, 回傳受影響筆數
        /// </summary>
        /// <returns>受影響筆數</returns>
        public int SaveChanges()
        {
            return context.SaveChanges();
        }

        /// <summary>
        /// 取得資料庫現在日期時間, 僅限MSSQL
        /// </summary>
        /// <returns>現在日期時間</returns>
        public DateTime GetSQLDateTime()
        {
            return SqlQuery<DateTime>("select getdate()").First();
        }
        /// 開始Transaction
        /// </summary>
        /// <param name="isolationLevel">transaction層級, 預設為Unspecified</param>
        public void BeginTransaction(IsolationLevel isolationLevel = IsolationLevel.Unspecified)
        {
            if (ObjContext.Connection.State != ConnectionState.Open)
            {
                ObjContext.Connection.Open();
            }

            _transaction = ObjContext.Connection.BeginTransaction(isolationLevel);
            logger.Debug("Begin Transaction");
        }

        /// <summary>
        /// commit transaction
        /// </summary>
        public void Commit()
        {
            _transaction.Commit();
            logger.Debug("Transaction Commit");
        }

        /// <summary>
        /// rollback transaction
        /// </summary>
        public void Rollback()
        {
            _transaction.Rollback();
            logger.Debug("Transaction Rollback");
        }

        /// <summary>
        /// 目前disposed狀態
        /// </summary>
        private bool disposed = false;

        /// <summary>
        /// 實作釋放資源
        /// </summary>
        /// <param name="disposing">是否釋放</param>
        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {
                    context.Dispose();
                }
            }
            this.disposed = true;
        }

        /// <summary>
        /// 釋放資源
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
    }
}
