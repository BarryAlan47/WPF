using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyApp
{
    class AaConvert
    {
        /// <summary>
        /// 人民币大小写转换函数
        /// </summary>
        /// <param name="amountStr"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public static string a2Afunc(string amountStr)
        {
            // 定义中文数字和单位
            string[] chineseDigits = { "零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖" };
            string[] units = { "", "拾", "佰", "仟" };
            string[] bigUnits = { "", "万", "亿" };
            string[] fractionalUnits = { "角", "分" };

            // 验证输入格式
            if (!decimal.TryParse(amountStr, out decimal amount) || amount < 0)
            {
                throw new ArgumentException("输入的金额格式无效");
            }

            // 将整数部分和小数部分分开
            string[] parts = amount.ToString("F2").Split('.');
            string integerPart = parts[0];
            string fractionalPart = parts[1];

            // 处理整数部分
            StringBuilder result = new StringBuilder("人民币");
            if (integerPart == "0")
            {
                result.Append("零元");
            }
            else
            {
                int bigUnitIndex = 0;
                bool hasZero = false;

                for (int i = integerPart.Length; i > 0; i -= 4)
                {
                    int length = Math.Min(4, i);
                    string subPart = integerPart.Substring(i - length, length);
                    StringBuilder subResult = new StringBuilder();

                    for (int j = 0; j < subPart.Length; j++)
                    {
                        int digit = subPart[j] - '0';
                        if (digit != 0)
                        {
                            if (hasZero)
                            {
                                subResult.Append(chineseDigits[0]);
                                hasZero = false;
                            }
                            subResult.Append(chineseDigits[digit]).Append(units[subPart.Length - j - 1]);
                        }
                        else
                        {
                            hasZero = true;
                        }
                    }

                    if (subResult.Length > 0)
                    {
                        subResult.Append(bigUnits[bigUnitIndex]);
                    }

                    result.Insert(3, subResult);
                    bigUnitIndex++;
                }

                result.Append("元");
            }

            // 处理小数部分
            if (fractionalPart == "00")
            {
                result.Append("整");
            }
            else
            {
                for (int i = 0; i < fractionalPart.Length; i++)
                {
                    int digit = fractionalPart[i] - '0';
                    if (digit != 0)
                    {
                        result.Append(chineseDigits[digit]).Append(fractionalUnits[i]);
                    }
                }
            }

            return result.ToString();
        }
    }
}
