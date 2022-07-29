using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ganss.Excel
{
    /// <summary>
    /// A caching factory of <see cref="TypeMapper"/> objects.
    /// </summary>
    public class DataColumnMapperFactory : ITypeMapperFactory
    {
        Dictionary<Type, TypeMapper> TypeMappers { get; set; } = new Dictionary<Type, TypeMapper>();

        /// <summary>
        /// Creates a <see cref="TypeMapper"/> for the specified type.
        /// </summary>
        /// <param name="type">The type to create a <see cref="TypeMapper"/> object for.</param>
        /// <returns>A <see cref="TypeMapper"/> for the specified type.</returns>
        public TypeMapper Create(Type type)
        {
            if (!TypeMappers.TryGetValue(type, out TypeMapper typeMapper))
            {
                typeMapper = TypeMappers[type] = TypeMapper.Create(type);
                typeMapper.ColumnsByIndex.Clear();
            }

            return typeMapper;
        }

        /// <summary>
        /// Creates a <see cref="TypeMapper"/> for the specified object.
        /// </summary>
        /// <param name="o">The object to create a <see cref="TypeMapper"/> object for.</param>
        /// <returns>A <see cref="TypeMapper"/> for the specified object.</returns>
        public TypeMapper Create(object o)
        {
            if (o is ExpandoObject eo)
            {
                if (!TypeMappers.TryGetValue(eo.GetType(), out TypeMapper typeMapper))
                {
                    typeMapper = TypeMappers[eo.GetType()] = TypeMapper.Create(eo);
                    typeMapper.ColumnsByIndex.Clear();
                }
                
                return typeMapper;
            }
            else
            {
                return Create(o.GetType());
            }
        }
    }
}
