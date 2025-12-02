using System;
using System.Collections.Generic;
using System.Text;

namespace MyOffice
{
    /// <summary>
    /// 统一结果返回类
    /// </summary>
    public class Result
    {
        /// <summary>
        /// 是否成功
        /// </summary>
        public bool Success { get; set; }
        
        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }
        
        /// <summary>
        /// 成功结果
        /// </summary>
        public static Result Ok()
        {
            return new Result { Success = true };
        }
        
        /// <summary>
        /// 失败结果
        /// </summary>
        public static Result Fail(string errorMessage)
        {
            return new Result { Success = false, ErrorMessage = errorMessage };
        }
    }
    
    /// <summary>
    /// 带数据的统一结果返回类
    /// </summary>
    /// <typeparam name="T">数据类型</typeparam>
    public class Result<T> : Result
    {
        /// <summary>
        /// 数据
        /// </summary>
        public T Data { get; set; }
        
        /// <summary>
        /// 成功结果
        /// </summary>
        public static Result<T> Ok(T data)
        {
            return new Result<T> { Success = true, Data = data };
        }
        
        /// <summary>
        /// 失败结果
        /// </summary>
        public static Result<T> Fail(string errorMessage)
        {
            return new Result<T> { Success = false, ErrorMessage = errorMessage };
        }
    }
}