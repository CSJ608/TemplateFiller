namespace TemplateFiller.Consts
{
    /// <summary>
    /// 模板占位符
    /// </summary>
    public static class PlaceholderConsts
    {
        /// <summary>
        /// 简单值占位符。使用":"分隔不同层级。
        /// </summary>
        /// <remarks>
        /// <example>
        /// <code>
        /// var p1 = "{Name}";                      // 在占位符处填充一个值，值来源为 Source.Name
        /// var p2 = "{Customer:Name}";             // 在占位符处填充一个值，值来源为 Source.Customer.Name
        /// var p3 = "{Project:Customer:Name}";     // 在占位符处填充一个值，值来源为 Source.Project.Customer.Name
        /// </code>
        /// </example>
        /// </remarks>
        public const string ValuePlaceholder = @"\{([a-zA-Z0-9:]+)\}";
        /// <summary>
        /// 数组占位符。使用":"分隔不同层级，并且允许使用一次"."，声明可枚举集合每个元素应取哪个字段
        /// </summary>
        /// <remarks>
        /// <example>
        /// <code>
        /// var p1 = "[Names]";                     // 在占位符及其下方，填充多个值，值来源为 Source.Names
        /// var p2 = "[Student:Names]";             // 在占位符及其下方，填充多个值，值来源为 Source.Student.Names
        /// var p3 = "[Students.Name]";             // 在占位符及其下方，填充多个值，值来源为 Source.Students.Select(item => item.Name)
        /// var p4 = "[Students.Info:Name]";        // 在占位符及其下方，填充多个值，值来源为 Source.Students.Select(item => item.Info.Name)
        /// var p5 = "[Project:Students.Name]";     // 在占位符及其下方，填充多个值，值来源为 Source.Project.Students.Select(item => item.Name)
        /// var p6 = "[Project:Students.Info:Name]";// 在占位符及其下方，填充多个值，值来源为 Source.Project.Students.Select(item => item.Info.Name)
        /// </code>
        /// </example>
        /// </remarks>
        public const string ArrayPlaceholder = @"\[(?<collectionPath>[a-zA-Z0-9]+(?::[a-zA-Z0-9]+)*)(?:\.(?<propertyPath>[a-zA-Z0-9:]+(?:\.[a-zA-Z0-9:]+)*))?\]";
    }
}
