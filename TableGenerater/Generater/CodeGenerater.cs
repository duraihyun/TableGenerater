using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.CodeDom;
using System.CodeDom.Compiler;
using Microsoft.CSharp;
using System.IO;

namespace Generater
{
    /// <summary>
    /// 코드 문서 객체 모델을 이용하여 코드를 생성한다.
    /// </summary>
    static class CodeGenerater
    {
        /// <summary>
        /// 스트링 -> type 컨테이너
        /// </summary>
        private static readonly Dictionary<string, Type> innerTypes = new Dictionary<string, Type>();

        /// <summary>
        /// 언어 스펙에서 지원하는 타입의 스트링 표현 컨테이너
        /// </summary>
        private static void InitializeTypes()
        {
            innerTypes.Add("sbyte", typeof(SByte));
            innerTypes.Add("int8", typeof(SByte));
            innerTypes.Add("byte", typeof(Byte));
            innerTypes.Add("uint8", typeof(Byte));
            innerTypes.Add("short", typeof(Int16));
            innerTypes.Add("int16", typeof(Int16));
            innerTypes.Add("ushort", typeof(UInt16));
            innerTypes.Add("uint16", typeof(UInt16));
            innerTypes.Add("int", typeof(Int32));
            innerTypes.Add("int32", typeof(Int32));
            innerTypes.Add("uint", typeof(UInt32));
            innerTypes.Add("uint32", typeof(UInt32));
            innerTypes.Add("long", typeof(Int64));
            innerTypes.Add("int64", typeof(Int64));
            innerTypes.Add("ulong", typeof(UInt64));
            innerTypes.Add("uint64", typeof(UInt64));
            innerTypes.Add("float", typeof(Single));
            innerTypes.Add("double", typeof(Double));
            innerTypes.Add("bool", typeof(Boolean));
            innerTypes.Add("boolean", typeof(Boolean));
            innerTypes.Add("string", typeof(String));
            innerTypes.Add("datetime", typeof(DateTime));
        }

        /// <summary>
        /// 타입명을 Type으로 변경한다.
        /// </summary>
        /// <param name="typeName">타입명</param>
        /// <returns>실패 시 null</returns>
        public static Type ConvertToTypeOrNull(string typeName)
        {
            if (false == innerTypes.Any())
            {
                InitializeTypes();
            }

            Type temp;
            if (true == innerTypes.TryGetValue(typeName.ToLower(), out temp))
            {
                return temp;
            }

            return null;
        }


        public static object ConvertToTypeValue(Type type, string value)
        {
            if (false == value.Any())
            {
                return string.Empty;
            }

            if (null != type)
            {
                return Convert.ChangeType(value, type);
            }

            // 내장 열거형 타입 확인
            return Enum.Parse(type, value);
        }



        /// <summary>
        /// 코드 네임스페이스를 이용하여 소스 코드를 생성한다.
        /// </summary>
        /// <param name="filePath">파일 경로</param>
        /// <param name="codeNamespace">코드로 변환할 코드 네임스페이스</param>
        public static void CrateCodeFile(string filePath, CodeNamespace codeNamespace)
        {
            var codeOptions = new CodeGeneratorOptions
            {
                BlankLinesBetweenMembers = false,
                VerbatimOrder = true,
                BracingStyle = "C",
                IndentString = "\t"
            };

            using (TextWriter tw = new StreamWriter(filePath, false, Encoding.GetEncoding(65001)))
            {
                using (var codeProvider = new CSharpCodeProvider())
                {
                    codeProvider.GenerateCodeFromNamespace(codeNamespace, tw, codeOptions);
                }
            }

            File.WriteAllText(filePath, File.ReadAllText(filePath).Replace("//;", ""), Encoding.GetEncoding(65001));
        }
    }
}
