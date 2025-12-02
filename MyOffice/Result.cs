using System;

namespace MyOffice
{
    /// <summary>
    /// 通用返回结果类
    /// </summary>
    /// <typeparam name="T">结果数据类型</typeparam>
    public class Result<T>
    {
        /// <summary>
        /// 是否成功
        /// </summary>
        public bool Success { get; set; }
        
        /// <summary>
        /// 错误消息
        /// </summary>
        public string ErrorMessage { get; set; }
        
        /// <summary>
        /// 结果数据
        /// </summary>
        public T Data { get; set; }
        
        /// <summary>
        /// 创建成功结果
        /// </summary>
        public static Result<T> Ok(T data)
        {
            return new Result<T> { Success = true, Data = data };
        }
        
        /// <summary>
        /// 创建失败结果
        /// </summary>
        public static Result<T> Fail(string errorMessage)
        {
            return new Result<T> { Success = false, ErrorMessage = errorMessage };
        }
    }
    
    /// <summary>
    /// 无数据返回结果类
    /// </summary>
    public class Result
    {
        /// <summary>
        /// 是否成功
        /// </summary>
        public bool Success { get; set; }
        
        /// <summary>
        /// 错误消息
        /// </summary>
        public string ErrorMessage { get; set; }
        
        /// <summary>
        /// 创建成功结果
        /// </summary>
        public static Result Ok()
        {
            return new Result { Success = true };
        }
        
        /// <summary>
        /// 创建失败结果
        /// </summary>
        public static Result Fail(string errorMessage)
        {
            return new Result { Success = false, ErrorMessage = errorMessage };
        }
    }
}