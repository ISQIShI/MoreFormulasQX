using System;
using System.Reflection;
using UnityEngine;

namespace MoreFormulasQX.Utils
{
    /// <summary>
    /// 反射辅助类，用于简化和优化反射操作
    /// </summary>
    public static class ReflectionHelper
    {
        private const BindingFlags DefaultBindingFlags =
            BindingFlags.Public |
            BindingFlags.NonPublic |
            BindingFlags.Instance |
            BindingFlags.Static;

        #region 字段操作

        /// <summary>
        /// 获取字段信息
        /// </summary>
        public static FieldInfo GetFieldInfo(Type type, string fieldName, BindingFlags? bindingFlags = null)
        {
            var flags = bindingFlags ?? DefaultBindingFlags;
            var field = type.GetField(fieldName, flags);

            if (field == null)
            {
                // 尝试在基类中查找
                Type baseType = type.BaseType;
                while (baseType != null)
                {
                    field = baseType.GetField(fieldName, flags);
                    if (field != null) break;
                    baseType = baseType.BaseType;
                }
            }
            return field;
        }

        /// <summary>
        /// 获取字段值
        /// </summary>
        public static T GetFieldValue<T>(object obj, string fieldName, BindingFlags? bindingFlags = null)
        {
            if (obj == null)
            {
                Debug.LogError("GetFieldValue: 目标对象为 null");
                return default;
            }

            try
            {
                var field = GetFieldInfo(obj.GetType(), fieldName, bindingFlags);

                if (field == null)
                {
                    Debug.LogError($"GetFieldValue: 无法找到字段 '{fieldName}' 在类型 '{obj.GetType().Name}' 中");
                    return default;
                }

                object value = field.GetValue(obj);

                if (value == null)
                {
                    return default;
                }

                if (value is T typedValue)
                {
                    return typedValue;
                }

                // 尝试转换
                try
                {
                    return (T)Convert.ChangeType(value, typeof(T));
                }
                catch
                {
                    Debug.LogError($"GetFieldValue: 无法将字段 '{fieldName}' 的值从 '{value.GetType().Name}' 转换为 '{typeof(T).Name}'");
                    return default;
                }
            }
            catch (Exception e)
            {
                Debug.LogError($"GetFieldValue: 获取字段 '{fieldName}' 时发生异常: {e.Message}");
                return default;
            }
        }

        /// <summary>
        /// 获取字段值（非泛型版本）
        /// </summary>
        public static object GetFieldValue(object obj, string fieldName, BindingFlags? bindingFlags = null)
        {
            if (obj == null)
            {
                Debug.LogError("GetFieldValue: 目标对象为 null");
                return null;
            }

            try
            {
                var field = GetFieldInfo(obj.GetType(), fieldName, bindingFlags);

                if (field == null)
                {
                    Debug.LogError($"GetFieldValue: 无法找到字段 '{fieldName}' 在类型 '{obj.GetType().Name}' 中");
                    return null;
                }

                return field.GetValue(obj);
            }
            catch (Exception e)
            {
                Debug.LogError($"GetFieldValue: 获取字段 '{fieldName}' 时发生异常: {e.Message}");
                return null;
            }
        }

        /// <summary>
        /// 设置字段值
        /// </summary>
        public static bool SetFieldValue(object obj, string fieldName, object value, BindingFlags? bindingFlags = null)
        {
            if (obj == null)
            {
                Debug.LogError("SetFieldValue: 目标对象为 null");
                return false;
            }

            try
            {
                var field = GetFieldInfo(obj.GetType(), fieldName, bindingFlags);

                if (field == null)
                {
                    Debug.LogError($"SetFieldValue: 无法找到字段 '{fieldName}' 在类型 '{obj.GetType().Name}' 中");
                    return false;
                }

                // 类型检查
                if (value != null && !field.FieldType.IsAssignableFrom(value.GetType()))
                {
                    // 尝试转换
                    try
                    {
                        value = Convert.ChangeType(value, field.FieldType);
                    }
                    catch
                    {
                        Debug.LogError($"SetFieldValue: 无法将值从 '{value.GetType().Name}' 转换为字段 '{fieldName}' 的类型 '{field.FieldType.Name}'");
                        return false;
                    }
                }

                field.SetValue(obj, value);
                return true;
            }
            catch (Exception e)
            {
                Debug.LogError($"SetFieldValue: 设置字段 '{fieldName}' 时发生异常: {e.Message}");
                return false;
            }
        }

        #endregion

        #region 属性操作

        /// <summary>
        /// 获取属性信息
        /// </summary>
        public static PropertyInfo GetPropertyInfo(Type type, string propertyName, BindingFlags? bindingFlags = null)
        {
            var flags = bindingFlags ?? DefaultBindingFlags;
            var property = type.GetProperty(propertyName, flags);

            if (property == null)
            {
                // 尝试在基类中查找
                Type baseType = type.BaseType;
                while (baseType != null)
                {
                    property = baseType.GetProperty(propertyName, flags);
                    if (property != null) break;
                    baseType = baseType.BaseType;
                }
            }

            return property;
        }

        /// <summary>
        /// 获取属性值
        /// </summary>
        public static T GetPropertyValue<T>(object obj, string propertyName, BindingFlags? bindingFlags = null)
        {
            if (obj == null)
            {
                Debug.LogError("GetPropertyValue: 目标对象为 null");
                return default;
            }

            try
            {
                var property = GetPropertyInfo(obj.GetType(), propertyName, bindingFlags);

                if (property == null)
                {
                    Debug.LogError($"GetPropertyValue: 无法找到属性 '{propertyName}' 在类型 '{obj.GetType().Name}' 中");
                    return default;
                }

                if (!property.CanRead)
                {
                    Debug.LogError($"GetPropertyValue: 属性 '{propertyName}' 不可读");
                    return default;
                }

                object value = property.GetValue(obj);

                if (value == null)
                {
                    return default;
                }

                if (value is T typedValue)
                {
                    return typedValue;
                }

                return (T)Convert.ChangeType(value, typeof(T));
            }
            catch (Exception e)
            {
                Debug.LogError($"GetPropertyValue: 获取属性 '{propertyName}' 时发生异常: {e.Message}");
                return default;
            }
        }

        /// <summary>
        /// 设置属性值
        /// </summary>
        public static bool SetPropertyValue(object obj, string propertyName, object value, BindingFlags? bindingFlags = null)
        {
            if (obj == null)
            {
                Debug.LogError("SetPropertyValue: 目标对象为 null");
                return false;
            }

            try
            {
                var property = GetPropertyInfo(obj.GetType(), propertyName, bindingFlags);

                if (property == null)
                {
                    Debug.LogError($"SetPropertyValue: 无法找到属性 '{propertyName}' 在类型 '{obj.GetType().Name}' 中");
                    return false;
                }

                if (!property.CanWrite)
                {
                    Debug.LogError($"SetPropertyValue: 属性 '{propertyName}' 不可写");
                    return false;
                }

                // 类型检查
                if (value != null && !property.PropertyType.IsAssignableFrom(value.GetType()))
                {
                    try
                    {
                        value = Convert.ChangeType(value, property.PropertyType);
                    }
                    catch
                    {
                        Debug.LogError($"SetPropertyValue: 无法将值转换为属性 '{propertyName}' 的类型");
                        return false;
                    }
                }

                property.SetValue(obj, value);
                return true;
            }
            catch (Exception e)
            {
                Debug.LogError($"SetPropertyValue: 设置属性 '{propertyName}' 时发生异常: {e.Message}");
                return false;
            }
        }

        #endregion

        #region 方法调用

        /// <summary>
        /// 获取方法信息
        /// </summary>
        public static MethodInfo GetMethodInfo(Type type, string methodName, Type[] parameterTypes = null, BindingFlags? bindingFlags = null)
        {
            var flags = bindingFlags ?? DefaultBindingFlags;
            MethodInfo method;

            if (parameterTypes != null)
            {
                method = type.GetMethod(methodName, flags, null, parameterTypes, null);
            }
            else
            {
                method = type.GetMethod(methodName, flags);
            }

            if (method == null)
            {
                // 尝试在基类中查找
                Type baseType = type.BaseType;
                while (baseType != null)
                {
                    if (parameterTypes != null)
                    {
                        method = baseType.GetMethod(methodName, flags, null, parameterTypes, null);
                    }
                    else
                    {
                        method = baseType.GetMethod(methodName, flags);
                    }

                    if (method != null) break;
                    baseType = baseType.BaseType;
                }
            }

            return method;
        }

        /// <summary>
        /// 调用方法
        /// </summary>
        public static object InvokeMethod(object obj, string methodName, object[] parameters = null, BindingFlags? bindingFlags = null)
        {
            if (obj == null)
            {
                Debug.LogError("InvokeMethod: 目标对象为 null");
                return null;
            }

            try
            {
                Type[] parameterTypes = null;
                if (parameters != null && parameters.Length > 0)
                {
                    parameterTypes = Array.ConvertAll(parameters, p => p?.GetType() ?? typeof(object));
                }

                var method = GetMethodInfo(obj.GetType(), methodName, parameterTypes, bindingFlags);

                if (method == null)
                {
                    Debug.LogError($"InvokeMethod: 无法找到方法 '{methodName}' 在类型 '{obj.GetType().Name}' 中");
                    return null;
                }

                return method.Invoke(obj, parameters);
            }
            catch (Exception e)
            {
                Debug.LogError($"InvokeMethod: 调用方法 '{methodName}' 时发生异常: {e.Message}\n{e.StackTrace}");
                return null;
            }
        }

        /// <summary>
        /// 调用方法（泛型返回值版本）
        /// </summary>
        public static T InvokeMethod<T>(object obj, string methodName, object[] parameters = null, BindingFlags? bindingFlags = null)
        {
            object result = InvokeMethod(obj, methodName, parameters, bindingFlags);

            if (result == null)
            {
                return default;
            }

            try
            {
                return (T)result;
            }
            catch
            {
                Debug.LogError($"InvokeMethod: 无法将返回值从 '{result.GetType().Name}' 转换为 '{typeof(T).Name}'");
                return default;
            }
        }

        #endregion

        #region 工具方法

        /// <summary>
        /// 检查字段是否存在
        /// </summary>
        public static bool HasField(Type type, string fieldName, BindingFlags? bindingFlags = null)
        {
            return GetFieldInfo(type, fieldName, bindingFlags) != null;
        }

        /// <summary>
        /// 检查属性是否存在
        /// </summary>
        public static bool HasProperty(Type type, string propertyName, BindingFlags? bindingFlags = null)
        {
            return GetPropertyInfo(type, propertyName, bindingFlags) != null;
        }

        /// <summary>
        /// 检查方法是否存在
        /// </summary>
        public static bool HasMethod(Type type, string methodName, Type[] parameterTypes = null, BindingFlags? bindingFlags = null)
        {
            return GetMethodInfo(type, methodName, parameterTypes, bindingFlags) != null;
        }

        #endregion
    }
}