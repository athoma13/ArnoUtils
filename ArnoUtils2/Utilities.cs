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
    public enum GenerationType
    {
        GetterOnly,
        GetterAndSetter,
        NotifyableGetterAndSetter,
    }

    [Flags()]
    public enum ParentType
    {
        vbClass = 1,
        vbStruct = 2,
        vbInterface = 4,
        vbEnum = 8,
    }

    public static class Utilities
    {

        private static TextSelection GetTextSelection(DTE2 application)
        {
            return application.ActiveDocument.Selection as TextSelection;
        }

        private static string GetPropertyName(string fieldName)
        {
            var tmp = fieldName.TrimStart('_');
            return tmp.Substring(0, 1).ToUpperInvariant() + tmp.Substring(1);
        }

        private static string GetLocalVariableName(string fieldName)
        {
            var tmp = fieldName.TrimStart('_');
            return tmp.Substring(0, 1).ToLowerInvariant() + tmp.Substring(1);
        }


        private static string GetFieldName(string parameterName)
        {
            return "_" + parameterName;
        }

        private static CodeElement GetParentSafe(TextPoint point, vsCMElement elementType)
        {
            try
            {
                return point.CodeElement[elementType];
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string[] GetUsingStatements(DTE2 application)
        {
            return application.ActiveDocument.ProjectItem.FileCodeModel.CodeElements.OfType<CodeImport>()
                            .Select(m => m.Namespace)
                            .ToArray();
        }

        private static CodeElement[] WhatsMyParent(TextPoint point)
        {
            var result = new List<CodeElement>();
            foreach (var value in Enum.GetValues(typeof(vsCMElement)).OfType<vsCMElement>())
            {
                var parent = GetParentSafe(point, value);
                if (parent != null) result.Add(parent);
            }
            return result.ToArray();
        }



        //Returns a CodeClass2 or CodeStruct2 object
        private static CodeElement GetParent(TextPoint point, ParentType parentType = ParentType.vbClass | ParentType.vbStruct)
        {

            CodeElement result = null;

            CodeElement parentClass = null;
            CodeElement parentStruct = null;
            CodeElement parentInterface = null;
            CodeElement parentEnum = null;

            if (((parentType & ParentType.vbClass) == ParentType.vbClass)) parentClass = GetParentSafe(point, vsCMElement.vsCMElementClass);
            if (((parentType & ParentType.vbStruct) == ParentType.vbStruct)) parentStruct = GetParentSafe(point, vsCMElement.vsCMElementStruct);
            if (((parentType & ParentType.vbInterface) == ParentType.vbInterface)) parentInterface = GetParentSafe(point, vsCMElement.vsCMElementInterface);
            if (((parentType & ParentType.vbEnum) == ParentType.vbEnum)) parentEnum = GetParentSafe(point, vsCMElement.vsCMElementEnum);


            var codeElements = new CodeElement[] { parentClass, parentStruct, parentInterface, parentEnum };

            foreach (var codeElement in codeElements)
            {
                if (codeElement != null)
                {
                    if (result == null)
                    {
                        result = codeElement;
                    }
                    else
                    {
                        if ((codeElement.StartPoint.AbsoluteCharOffset > result.StartPoint.AbsoluteCharOffset))
                        {
                            result = codeElement;
                        }
                    }
                }

            }


            if (result == null)
            {
                throw new InvalidOperationException("Must be within " + parentType.ToString());
            }


            return result;
        }

        private static CodeElement2[] GetCodeElementsInSelection(TextSelection selection, params vsCMElement[] kinds)
        {
            var result = new List<CodeElement2>();
            int stPoint;
            int enPoint;


            //If no selection is made, select the whole line
            if (selection.TopPoint.AbsoluteCharOffset == selection.BottomPoint.AbsoluteCharOffset)
            {
                //'Select the whole line
                stPoint = selection.TopPoint.AbsoluteCharOffset - selection.TopPoint.LineCharOffset;
                enPoint = stPoint + selection.TopPoint.LineLength + 1;
            }
            else
            {
                stPoint = selection.TopPoint.AbsoluteCharOffset;
                enPoint = selection.BottomPoint.AbsoluteCharOffset;
            }


            foreach (var element in GetParent(selection.TopPoint).Children.OfType<CodeElement2>())
            {
                if (kinds.Contains(element.Kind)
                        && element.StartPoint.AbsoluteCharOffset >= stPoint
                        && element.EndPoint.AbsoluteCharOffset <= enPoint)
                {
                    result.Add(element);
                }
            }

            return result.OrderBy(c => c.StartPoint.AbsoluteCharOffset).ToArray();
        }


        private static CodeElement FindElement(string name, CodeElements elements, vsCMElement kind)
        {
            return FindElement(name, elements.OfType<CodeElement>(), kind);
        }
        private static CodeElement FindElement(string name, IEnumerable<CodeElement> elements, vsCMElement kind)
        {
            foreach (var element in elements)
            {
                if (element.Name == name && element.Kind == kind)
                {
                    return element;
                }
            }
            return null;
        }

        private static CodeElement FindLast(vsCMElement kind, CodeElement parent)
        {
            CodeElement result = null;
            foreach (CodeElement element in parent.Children)
            {
                if (element.Kind == kind)
                {
                    result = element;
                }
            }
            return result;
        }

        private static string GetBody(SourceElement sourceElement, GenerationType genType)
        {
            var sb = new StringBuilder();
            var sw = new StringWriter(sb);

            if (genType == GenerationType.GetterAndSetter)
            {
                sw.Write("get; set;");
            }
            else if (genType == GenerationType.GetterOnly)
            {
                sw.Write("get; private set;");
            }
            else if (genType == GenerationType.NotifyableGetterAndSetter)
            {
                sw.WriteLine("get {{ return {0}; }}", sourceElement.FieldName);
                sw.Write("set {1}{{{1} if ({0} != value){1} {{{1}{0} = value;{1}OnPropertyChanged(\"{2}\");{1}}}{1}}}{1}", sourceElement.FieldName, Environment.NewLine, sourceElement.PropertyName);
            }

            return sb.ToString();
        }


        private static void CreateOrReplaceProperty(SourceElement sourceElement, GenerationType genType, CodeElement parent)
        {
            var property = (CodeProperty)FindElement(sourceElement.PropertyName, parent.Children, vsCMElement.vsCMElementProperty);
            if (property == null)
            {
                //Create the property...

                var insertAfterElement = FindLast(vsCMElement.vsCMElementProperty, parent) ?? FindLast(vsCMElement.vsCMElementVariable, parent);
                var newPropertyPoint = insertAfterElement.EndPoint.CreateEditPoint();

                newPropertyPoint.Insert(Environment.NewLine);
                newPropertyPoint.Insert(string.Format("public {0} {1} {{get; set;}}", sourceElement.Type, sourceElement.PropertyName));
                property = (CodeProperty)FindElement(sourceElement.PropertyName, parent.Children, vsCMElement.vsCMElementProperty);
            }


            var field = (CodeVariable)FindElement(sourceElement.FieldName, parent.Children, vsCMElement.vsCMElementVariable);
            if (field != null)
            {
                if (property.Type.AsString != field.Type.AsString)
                {
                    property.Type = field.Type;
                }

                if (property.Getter.IsShared != field.IsShared)
                {
                    property.Getter.IsShared = field.IsShared;
                    property.Setter.IsShared = field.IsShared;
                }
            }

            if (genType == GenerationType.NotifyableGetterAndSetter && field == null)
            {
                var newFieldPoint = parent.GetStartPoint(vsCMPart.vsCMPartBody);
                var insertAfterElement = FindLast(vsCMElement.vsCMElementVariable, parent);
                if (insertAfterElement != null) newFieldPoint = insertAfterElement.EndPoint;
                newFieldPoint.CreateEditPoint().Insert(string.Format("\r\nprivate {0} {1};", sourceElement.Type, sourceElement.FieldName));
            }


            var bodyTxt = GetBody(sourceElement, genType);
            if (genType != GenerationType.NotifyableGetterAndSetter) DeleteField(sourceElement.FieldName, parent);

            var insertPoint = property.GetStartPoint(vsCMPart.vsCMPartBody).CreateEditPoint();
            insertPoint.Delete(property.GetEndPoint(vsCMPart.vsCMPartBody));
            insertPoint.Insert(bodyTxt);

            var formatPoint = parent.StartPoint.CreateEditPoint();
            formatPoint.SmartFormat(parent.EndPoint);
            //insertPoint.SmartFormat(property.StartPoint);

        }
        
        private static void DeleteField(string name, CodeElement parent)
        {
            var element = FindElement(name, parent.Children, vsCMElement.vsCMElementVariable);
            if (element == null) return;
            dynamic p = parent;
            p.RemoveMember(element);
        }

        private static void CreateGetterAndSetters(GenerationType genType, DTE2 application)
        {
            var selection = GetTextSelection(application);
            var parent = GetParent(selection.TopPoint);

            var sourceElements = GetCodeElementsInSelection(selection, vsCMElement.vsCMElementVariable, vsCMElement.vsCMElementProperty).Select(c => new SourceElement(c)).ToArray();

            foreach (var element in sourceElements)
            {
                CreateOrReplaceProperty(element, genType, parent);
            }
        }

        private static bool MakeNotifyable(DTE2 application)
        {
            var selection = GetTextSelection(application);
            var parent = GetParent(selection.TopPoint);
            var usings = GetUsingStatements(application);
            var ns = usings.Contains("System.ComponentModel") ? "" : "System.ComponentModel.";

            //Add the Interface
            var isNotifiable = false;
            //TODO: Dynamic?!?
            dynamic pp = parent;
            foreach (CodeElement interfaceElem in pp.ImplementedInterfaces)
            {
                if (interfaceElem.FullName == "System.ComponentModel.INotifyPropertyChanged")
                {
                    isNotifiable = true;
                    break;
                }
            }
            if (!isNotifiable)
            {
                pp.AddImplementedInterface(ns + "INotifyPropertyChanged");
            }


            //Add the Event
            var isPropertyChangedEventExists = false;

            foreach (CodeElement memberElem in pp.Members)
            {
                var ce = memberElem as CodeEvent;
                if (ce != null)
                {
                    if (ce.Access == vsCMAccess.vsCMAccessPublic && ce.Name == "PropertyChanged" && ce.Type.AsFullName == "System.ComponentModel.PropertyChangedEventHandler")
                    {
                        isPropertyChangedEventExists = true;
                        break;
                    }
                }
            }


            if (!isPropertyChangedEventExists)
            {
                pp.AddEvent("PropertyChanged", "System.ComponentModel.PropertyChangedEventHandler", null, null, vsCMAccess.vsCMAccessPublic);
            }




            //Add the Function
            var isOnPropertyChangedFunctionExists = false;
            foreach (CodeElement memberElem in pp.Members)
            {
                var mElem = memberElem as CodeFunction;
                if (mElem != null)
                {
                    if (!mElem.IsShared && mElem.Name == "OnPropertyChanged" && mElem.Parameters.Count == 1)
                    {
                        var isParaString = false;
                        foreach (CodeParameter mPara in mElem.Parameters)
                        {
                            if (mPara.Type.AsFullName == "System.String")
                            {
                                isParaString = true;
                            }
                        }
                        if (isParaString)
                        {
                            isOnPropertyChangedFunctionExists = true;
                            break;
                        }
                    }
                }
            }


            if (!isOnPropertyChangedFunctionExists)
            {
                var sb = new StringBuilder();

                sb.AppendLine();
                if (parent.Kind == vsCMElement.vsCMElementClass)
                    sb.AppendLine("protected virtual void OnPropertyChanged(string property)");
                else //'We are in a struct, no protected virtuals allowed!
                    sb.AppendLine("private void OnPropertyChanged(string property)");

                sb.AppendLine("{");
                sb.AppendLine("if (this.PropertyChanged != null)");
                sb.AppendLine("{");
                sb.AppendFormat("this.PropertyChanged(this, new {0}PropertyChangedEventArgs(property));", ns);
                sb.AppendLine();
                sb.AppendLine("}");
                sb.AppendLine("}");


                var pt = pp.EndPoint().CreateEditPoint();
                pt.MoveToAbsoluteOffset(pt.AbsoluteCharOffset - 1);
                var offset = pt.AbsoluteCharOffset;


                pt.Insert(sb.ToString());
                pt.MoveToAbsoluteOffset(offset);
                pt.SmartFormat(parent.EndPoint);
            }




            return true;
        }

        private static CodeParameter2[] GetSelectedParameters(TextSelection selection)
        {
            var result = new List<CodeParameter2>();
            var currentFunction = (CodeFunction)selection.TopPoint.CodeElement[vsCMElement.vsCMElementFunction];

            if (currentFunction != null)
            {
                var startOffset = selection.TopPoint.AbsoluteCharOffset;
                var bottomOffset = selection.BottomPoint.AbsoluteCharOffset;

                var tmp = selection.TopPoint.CodeElement[vsCMElement.vsCMElementParameter];
                if (tmp != null)
                {
                    startOffset = Math.Min(selection.TopPoint.AbsoluteCharOffset, tmp.StartPoint.AbsoluteCharOffset);
                }

                tmp = selection.BottomPoint.CodeElement[vsCMElement.vsCMElementParameter];
                if (tmp != null)
                {
                    bottomOffset = Math.Max(selection.BottomPoint.AbsoluteCharOffset, tmp.EndPoint.AbsoluteCharOffset);
                }

                foreach (CodeParameter2 param in currentFunction.Parameters)
                {
                    if (param.Kind == vsCMElement.vsCMElementParameter
                                && param.StartPoint.AbsoluteCharOffset >= startOffset
                                && param.EndPoint.AbsoluteCharOffset <= bottomOffset)
                    {
                        result.Add(param);
                    }
                }
            }
            return result.ToArray();
        }

        private static void CreateFieldFromConstructor(DTE2 application, CodeParameter2[] parameters)
        {
            dynamic parent = GetParent(parameters[0].StartPoint);
            var ctor = parameters[0].Parent as CodeFunction;

            foreach (CodeParameter2 parameter in parameters)
            {
                var fieldName = GetFieldName(parameter.Name);
                var tmp = string.Format("{0} = {1};" + Environment.NewLine, fieldName, parameter.Name);

                ctor.GetEndPoint(vsCMPart.vsCMPartBody).CreateEditPoint().Insert(tmp);

                if (FindElement(fieldName, parent.Children, vsCMElement.vsCMElementVariable) == null)
                {
                    //Create the member field
                    var lastField = FindLast(vsCMElement.vsCMElementVariable, parent) as CodeVariable2;
                    parent.AddVariable(fieldName, parameter.Type.AsString, lastField, vsCMAccess.vsCMAccessPrivate);
                }
            }
            ctor.StartPoint.CreateEditPoint().SmartFormat(ctor.EndPoint);
        }

        private static CodeFunction2 FindConstructor(CodeElement parent, string argumentName)
        {
            foreach (var element in parent.Children.OfType<CodeFunction2>())
            {
                if (element.Kind == vsCMElement.vsCMElementFunction && element.FunctionKind == vsCMFunction.vsCMFunctionConstructor)
                {
                    var result = element.Parameters.OfType<CodeElement2>().FirstOrDefault(p => p.Name == argumentName);
                    if (result == null) return element;
                }
            }
            return null;
        }




        private static EditPoint GetBodyStartPoint(CodeElement element)
        {
            var tmpPt = element.StartPoint.CreateEditPoint();
            if (tmpPt.FindPattern("{", (int)vsFindOptions.vsFindOptionsMatchInHiddenText, null, null))
            {
                tmpPt.CharRight();
                return tmpPt;
            }
            else
            {
                throw new InvalidOperationException("Cannot find '{' for code element " + element.Name);
            }
        }



        public static void CreateGetterAndSetters(DTE2 application)
        {
            CreateGetterAndSetters(GenerationType.GetterAndSetter, application);
        }


        public static void CreateGettersOnly(DTE2 application)
        {
            CreateGetterAndSetters(GenerationType.GetterOnly, application);
        }


        public static void CreateGetterAndSettersINotifiable(DTE2 application)
        {
            MakeNotifyable(application);
            CreateGetterAndSetters(GenerationType.NotifyableGetterAndSetter, application);
        }

        public static void GuardNull(DTE2 application)
        {
            var selection = GetTextSelection(application);
            var parameter = (CodeParameter)selection.ActivePoint.CodeElement[vsCMElement.vsCMElementParameter];

            if (parameter != null)
            {
                var usings = GetUsingStatements(application);
                var ns = usings.Contains("System") ? "" : "System.";

                var ed = parameter.Parent.GetStartPoint(vsCMPart.vsCMPartBody).CreateEditPoint();
                ed.Insert(String.Format("if ({0} == null) throw new {1}ArgumentNullException(\"{0}\");" + Environment.NewLine, parameter.Name, ns));
                parameter.Parent.GetStartPoint().CreateEditPoint().SmartFormat(parameter.Parent.GetEndPoint());
            }
        }

        public static void ConstructorMacro(DTE2 application)
        {
            //Creates a constructor or adds to an existing one OR if invoked from a constructor, creates a private member field
            Boolean isInConstructor = false;
            TextSelection selection = GetTextSelection(application);
            CodeParameter2[] parameters = GetSelectedParameters(selection);

            if (parameters.Length > 0
                && parameters[0].Parent.Kind == vsCMElement.vsCMElementFunction
                && ((CodeFunction)parameters[0].Parent).FunctionKind == vsCMFunction.vsCMFunctionConstructor
                )
            {
                isInConstructor = true;
            }


            if (isInConstructor)
            {
                CreateFieldFromConstructor(application, parameters);
            }
            else
            {
                CreateConstructorFromFields(selection);
            }
        }



        public static void MakeNewProjectItem(DTE2 application)
        {
            var selection = GetTextSelection(application);

            var parentElement = GetParent(selection.ActivePoint, ParentType.vbClass | ParentType.vbInterface | ParentType.vbStruct | ParentType.vbEnum);
            var nmspaceElement = parentElement.StartPoint.CodeElement[vsCMElement.vsCMElementNamespace];
            var innerNmspaceBodyStartPoint = GetBodyStartPoint(nmspaceElement);
            var classStartPoint = innerNmspaceBodyStartPoint;
            TextPoint classEndPoint = null;
            string classText = null;

            foreach (CodeElement2 child in nmspaceElement.Children)
            {
                if ((child.Name == parentElement.Name))
                {
                    classEndPoint = child.EndPoint;
                    classText = classStartPoint.GetText(classEndPoint).TrimStart('\r', '\n');
                    break;
                }
                else
                {
                    classStartPoint = child.EndPoint.CreateEditPoint();
                }
            }

            if (classText == null)
            {
                throw new InvalidOperationException("Cannot find class or structure " + parentElement.Name);
            }

            var docStartPoint = selection.ActivePoint.CreateEditPoint();
            docStartPoint.MoveToAbsoluteOffset(1);

            var usingAndNamespaceText = docStartPoint.GetText(innerNmspaceBodyStartPoint).Trim();
            var tempPath = Path.Combine(Path.GetTempPath(), parentElement.Name + ".cs");
            var text = new StringBuilder();

            text.AppendLine(usingAndNamespaceText);
            text.AppendLine(classText);
            text.AppendLine("}");

            File.WriteAllText(tempPath, text.ToString());

            parentElement.ProjectItem.Collection.AddFromFileCopy(tempPath);
            classStartPoint.Delete(classEndPoint);

        }


        //TODO: Order Getters and Setters in the same order as the fields!
        //TODO: ctrl+g,ctrl+g should rename the field if invoked from within a Getter/Setter (Should also see if the reference in the constructors could be renamed) - Convert To AutoProperty? Toggle (ctrl+g,ctrl+a)?


        private static void CreateConstructorFromFields(TextSelection selection)
        {
            dynamic parent = GetParent(selection.TopPoint);

            var sourceElements = GetCodeElementsInSelection(selection, vsCMElement.vsCMElementVariable, vsCMElement.vsCMElementProperty).Select(c => new SourceElement(c)).ToArray();
            var ctor = FindConstructor(parent, GetLocalVariableName(sourceElements[0].ParameterName));

            //TODO: Should not select the constructor when creating a new constructor (happens if there are only fields and no constructor).
            if (ctor == null)
            {
                var insertionPoint = FindLast(vsCMElement.vsCMElementProperty, parent);
                if (insertionPoint == null)
                {
                    insertionPoint = FindLast(vsCMElement.vsCMElementVariable, parent);
                }
                ctor = parent.AddFunction(parent.Name, vsCMFunction.vsCMFunctionConstructor, null, insertionPoint, vsCMAccess.vsCMAccessPublic);
            }
            var sb = new StringBuilder();
            foreach (var element in sourceElements)
            {
                if (FindElement(element.ParameterName, ctor.Parameters, vsCMElement.vsCMElementParameter) == null)
                {
                    sb.AppendFormat("{0} = {1};{2}", element.Name, element.ParameterName, Environment.NewLine);
                    ctor.AddParameter(element.ParameterName, element.Type, ctor.Parameters.Count);
                }
            }

            ctor.GetEndPoint(vsCMPart.vsCMPartBody).CreateEditPoint().Insert(sb.ToString());
            ctor.StartPoint.CreateEditPoint().SmartFormat(ctor.EndPoint);
        }
    }
}
