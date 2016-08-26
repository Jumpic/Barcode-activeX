using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Denvic.Barcode
{
    [ComVisible(true)]
    [Guid("AE2BDAA6-3A5B-4F81-8D9E-CE8903ECB48A")]
    public enum TypeBarcodeEnum
    {
        [Description("EAN_8")]
        EAN_8 = 0,
        [Description("EAN_13")]
        EAN_13 = 1,
        [Description("CODE_39")]
        CODE_39 = 2,
        [Description("CODE_128")]
        CODE_128 = 3,
        [Description("QR_CODE")]
        QR_CODE = 4,
        [Description("DATA_MATRIX")]
        DATA_MATRIX = 5
    }



    //class EnumTypeConverter : EnumConverter
    //{
    //    private Type _enumType;
    //    /// <summary>Инициализирует экземпляр</summary>
    //    /// <param name="type">тип Enum</param>
    //    public EnumTypeConverter(Type type) : base(type)
    //    {
    //        _enumType = type;
    //    }

    //    public override bool CanConvertTo(ITypeDescriptorContext context,
    //      Type destType)
    //    {
    //        return destType == typeof(string);
    //    }

    //    public override object ConvertTo(ITypeDescriptorContext context,
    //      CultureInfo culture,
    //      object value, Type destType)
    //    {
    //        FieldInfo fi = _enumType.GetField(Enum.GetName(_enumType, value));
    //        DescriptionAttribute dna =
    //          (DescriptionAttribute)Attribute.GetCustomAttribute(
    //            fi, typeof(DescriptionAttribute));

    //        if (dna != null)
    //            return dna.Description;
    //        else
    //            return value.ToString();
    //    }

    //    public override bool CanConvertFrom(ITypeDescriptorContext context,
    //      Type srcType)
    //    {
    //        return srcType == typeof(string);
    //    }

    //    public override object ConvertFrom(ITypeDescriptorContext context,
    //      CultureInfo culture,
    //      object value)
    //    {
    //        foreach (FieldInfo fi in _enumType.GetFields())
    //        {
    //            DescriptionAttribute dna =
    //              (DescriptionAttribute)Attribute.GetCustomAttribute(
    //                fi, typeof(DescriptionAttribute));

    //            if ((dna != null) && ((string)value == dna.Description))
    //                return Enum.Parse(_enumType, fi.Name);
    //        }

    //        return Enum.Parse(_enumType, (string)value);
    //    }

    //}
}
