namespace WorkbookSharp;

public static class TypeExtensions
{
    public static bool IsTuple(this Type type) => type.IsGenericType && type.FullName!.StartsWith("System.ValueTuple");

    public static bool IsAnonymous(this Type type) => Attribute.IsDefined(type, typeof(System.Runtime.CompilerServices.CompilerGeneratedAttribute)) && 
                                                      type.IsGenericType &&
                                                      type.Name.Contains("AnonymousType") &&
                                                      (type.Name.StartsWith("<>") || type.Name.StartsWith("VB$")) &&
                                                      (type.Attributes & System.Reflection.TypeAttributes.NotPublic) == System.Reflection.TypeAttributes.NotPublic;
    
}
