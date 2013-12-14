using EnvDTE;
using EnvDTE80;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ArnoUtils2
{
    public class SourceElement
    {
        public const string PREFIX = "_";

        public CodeElement OriginalElement { get; private set; }
        public string Name { get; private set; }
        public bool IsProperty { get; private set; }
        public bool IsField { get; private set; }
        public bool IsParameter { get; private set; }


        public string Type
        {
            get
            {
                dynamic d = OriginalElement;
                string result = d.Type.AsString;
                return result.Split('.').Last();
            }
        }

        public string FieldName
        {
            get
            {
                if (IsField) return Name;
                if (IsParameter) return PREFIX + ToLower(Name);
                if (IsProperty) return PREFIX + ToLower(Name);
                throw new InvalidOperationException("Invalid SourceElement");
            }
        }

        public string PropertyName
        {
            get
            {
                if (IsField) return ToUpper(TrimPrefix(Name));
                if (IsParameter) return ToUpper(Name);
                if (IsProperty) return Name;
                throw new InvalidOperationException("Invalid SourceElement");
            }
        }

        public string ParameterName
        {
            get
            {
                if (IsField) return ToLower(TrimPrefix(Name));
                if (IsParameter) return Name;
                if (IsProperty) return ToLower(Name);
                throw new InvalidOperationException("Invalid SourceElement");
            }
        }

        public SourceElement(CodeElement element)
        {
            if (element == null) throw new ArgumentNullException("element");
            if (element.Kind == vsCMElement.vsCMElementProperty) IsProperty = true;
            if (element.Kind == vsCMElement.vsCMElementParameter) IsParameter = true;
            if (element.Kind == vsCMElement.vsCMElementVariable) IsField = true;
            Name = element.Name;
            OriginalElement = element;
        }

        public void Delete()
        {
            dynamic d = OriginalElement;
            d.Parent.RemoveMember(OriginalElement);
        }

        private static string TrimPrefix(string value)
        {
            if (value.StartsWith(PREFIX)) return value.Substring(PREFIX.Length);
            return value;
        }

        private static string ToUpper(string value)
        {
            return ChangeFirstCharCase(value, true);
        }

        private static string ToLower(string value)
        {
            return ChangeFirstCharCase(value, false);
        }

        private static string ChangeFirstCharCase(string value, bool toUpper)
        {
            var buffer = value.ToCharArray();
            if (toUpper) buffer[0] = char.ToUpperInvariant(buffer[0]);
            else buffer[0] = char.ToLowerInvariant(buffer[0]);

            return new string(buffer);
        }

        private static void ValidateName(string name)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentNullException();
            if (string.IsNullOrEmpty(name.Trim())) throw new ArgumentNullException();
            if (name == PREFIX) throw new ArgumentException("name cannot be " + PREFIX);
        }
    }
}
