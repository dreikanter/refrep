using System;
using System.Text;

namespace WordReplace.Extensions
{
    public static class ExceptionExtensions
    {
        /// <summary>
        /// Returns exception Message merged with inner exception Message
        /// </summary>
        public static string GetMessage(this Exception exception)
        {
            var message = new StringBuilder(exception.Message);
            var innerException = exception.InnerException;
            var depth = 0;

            while (innerException != null && depth++ < 10)
            {
				message.AppendFormat("{0}Inner exception: {1}",
					Environment.NewLine, innerException.Message.TrimEnd(new[] { '.' }));
				innerException = innerException.InnerException;
            }

            return message.ToString();
        }
    }
}