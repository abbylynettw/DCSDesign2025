using System;

namespace MyOffice
{
    /// <summary>
    /// 统一的返回结果类
    /// </summary>
    /// <typeparam name="T">返回数据的类型</typeparam>
    public class Result<T>
    {
        /// <summary>
        /// 操作是否成功
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 返回的数据
        /// </summary>
        public T Data { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 异常对象
        /// </summary>
        public Exception Exception { get; set; }

        /// <summary>
        /// 创建一个成功的结果
        /// </summary>
        /// <param name="data">返回的数据</param>
        /// <returns>成功的结果对象</returns>
        public static Result<T> SuccessResult(T data = default)
        {
            return new Result<T>
            {
                Success = true,
                Data = data,
                ErrorMessage = null,
                Exception = null
            };
        }

        /// <summary>
        /// 创建一个失败的结果
        /// </summary>
        /// <param name="errorMessage">错误信息</param>
        /// <param name="exception">异常对象</param>
        /// <returns>失败的结果对象</returns>
        public static Result<T> FailureResult(string errorMessage, Exception exception = null)
        {
            return new Result<T>
            {
                Success = false,
                Data = default,
                ErrorMessage = errorMessage,
                Exception = exception
            };
        }
    }

    /// <summary>
    /// 无返回数据的结果类
    /// </summary>
    public class Result : Result<object>
    {
        /// <summary>
        /// 创建一个成功的结果
        /// </summary>
        /// <returns>成功的结果对象</returns>
        public new static Result SuccessResult()
        {
            return new Result
            {
                Success = true,
                Data = null,
                ErrorMessage = null,
                Exception = null
            };
        }

        /// <summary>
        /// 创建一个失败的结果
        /// </summary>
        /// <param name="errorMessage">错误信息</param>
        /// <param name="exception">异常对象</param>
        /// <returns>失败的结果对象</returns>
        public new static Result FailureResult(string errorMessage, Exception exception = null)
        {
            return new Result
            {
                Success = false,
                Data = null,
                ErrorMessage = errorMessage,
                Exception = exception
            };
        }
    }
}
