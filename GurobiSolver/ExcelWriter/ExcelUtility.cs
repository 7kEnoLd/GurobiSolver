using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Security.Claims;
using System.Text;
using System.Threading.Tasks;
using static System.Console;

namespace ExcelWriter_Utility
{
    public static class ToModel
    {
        #region
        /// <summary>
        /// excel处理的datatable数据转为带有模型结构的泛型,表中的标题要与模型的属性名字一样
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <param name="dr"></param>
        /// <returns></returns>
        public static TModel DataRowToModel<TModel>(this DataRow dr) //泛型反射，适用于多个类型的方法，扩展方法加了个this，必须在静态类静态方法中
        {
            //获取TModel的实例化对象，等价于new对象那一步
            Type type = typeof(TModel);
            TModel md = (TModel)Activator.CreateInstance(type);

            //获取type的各个属性
            foreach (var prop in type.GetProperties())
            {
                try
                {
                    //数据类型转换
                    type = prop.PropertyType;
                    var pr = Convert.ChangeType(dr[prop.Name], type);
                    //填充数据
                    prop.SetValue(md, pr);
                }
                catch (FormatException)
                {
                    //目前测试了超出数据上限的和非法格式类型的
                    WriteLine("数据格式有问题，请检查数据！");
                }
                catch (Exception ex)
                {
                    WriteLine($"{ex.GetType()} 显示 {ex.Message}");
                }
            }

            return md;
        }
        #endregion

        #region
        /// <summary>
        /// list转datatable，用于excel的数据导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <returns></returns>
        public static DataTable ToDataTable<T>(this IList<T> data)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            DataTable dt = new DataTable();
            for (int i = 0; i < properties.Count; i++)
            {
                PropertyDescriptor property = properties[i];
                dt.Columns.Add(property.Name, property.PropertyType);
            }
            object[] values = new object[properties.Count];
            foreach (T item in data)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = properties[i].GetValue(item);
                }
                dt.Rows.Add(values);
            }
            return dt;
        }
        #endregion


        #region
        /// <summary>
        /// 返回list泛型的各个属性的集合,想要的属性可以用list[i][j]接收，i代表属性索引，j代表对应属性下的元素索引
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="values"></param>
        /// <returns></returns>
        public static List<List<object?>> ToGetList<T>(List<T> values)
        {
            List<T> coorlist = new List<T>();

            //PropertyInfo[] PropertyList = coorlist[0].GetType().GetProperties();
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            List<List<object?>> allData = new List<List<object?>>();
            for (int i = 0; i < properties.Count; i++)
            {
                string name = properties[i].Name;
                //var thisvalue = coorlist.Select(p => p.GetType().GetProperty(name).GetValue(p)).Distinct().ToList(); 删除重复数据
                var thisvalue = coorlist.Select(p => p.GetType().GetProperty(name).GetValue(p)).ToList();
                allData.Add(thisvalue);
            }
            return allData;

        }
        #endregion


        #region 记录程序运行时间
        //System.Diagnostics.Stopwatch stopwatch = new Stopwatch();
        //stopwatch.Start(); //  开始监视代码运行时间
        //    //  you code ....
        //    stopwatch.Stop(); //  停止监视
        //    TimeSpan timespan = stopwatch.Elapsed; //  获取当前实例测量得出的总时间
        //double hours = timespan.TotalHours; // 总小时
        //double minutes = timespan.TotalMinutes;  // 总分钟
        //double seconds = timespan.TotalSeconds;  //  总秒数
        //double milliseconds = timespan.TotalMilliseconds;  // 
        #endregion

    }
}