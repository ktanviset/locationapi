using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Threading.Tasks;

namespace locationapi.AppExtensions
{
    public static class SqlExtensions
    {
        public static T Read<T>(DbDataReader DataReader, string FieldName)
        {
            int FieldIndex;
            try { FieldIndex = DataReader.GetOrdinal(FieldName); }
            catch { return default(T); }

            if (DataReader.IsDBNull(FieldIndex))
            {
                return default(T);
            }
            else
            {
                object readData = DataReader.GetValue(FieldIndex);
                if (readData is T)
                {
                    return (T)readData;
                }
                else
                {
                    try
                    {
                        return (T)Convert.ChangeType(readData, typeof(T));
                    }
                    catch (InvalidCastException)
                    {
                        return default(T);
                    }
                }
            }
        }
    }
}
