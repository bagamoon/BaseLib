using AutoMapper;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LibCommon.Extension
{
    public static partial class Extensions
    {
        /// <summary>
        /// Deep Copy物件, 透過JSON.Net將物件序列化再反序列化來實作
        /// 注意循環參考屬性將被忽略
        /// </summary>
        /// <typeparam name="T">物件類型</typeparam>
        /// <param name="source">欲Deep Copy的物件</param>
        /// <returns>Deep Copy的新物件</returns>
        public static T DeepClone<T>(this T source)
        {
            if (Object.ReferenceEquals(source, null))
            {
                return default(T);
            }

            JsonSerializerSettings setting = new JsonSerializerSettings();
            setting.ReferenceLoopHandling = ReferenceLoopHandling.Ignore;

            return JsonConvert.DeserializeObject<T>(JsonConvert.SerializeObject(source, setting));
        }

        /// <summary>
        /// AutoMapper擴充方法, 可使target class重複mapping多個source class
        /// </summary>
        /// <typeparam name="TSource">source class</typeparam>
        /// <typeparam name="TDestination">target class</typeparam>
        /// <param name="destination">target class instance</param>
        /// <param name="source">source class instance</param>
        /// <returns>mapping完成的target class instance</returns>
        public static TDestination Map<TSource, TDestination>(this TDestination destination, TSource source)
        {
            if (source == null)
            {
                return destination;
            }

            return Mapper.Map(source, destination);
        }
    }
}
