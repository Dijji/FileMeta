using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Automation;
using AutomationExample;

// Taken from http://blog.functionalfun.net/2009/06/introduction-to-ui-automation-with.html
// No visible copyright

namespace AutomationExample
{
    public static class AutomationExtensions
    {
        public static void EnsureElementIsScrolledIntoView(this AutomationElement element)
        {
            if (!element.Current.IsOffscreen)
            {
                return;
            }

            if (!(bool)element.GetCurrentPropertyValue(AutomationElement.IsScrollItemPatternAvailableProperty))
            {
                return;
            }

            var scrollItemPattern = element.GetScrollItemPattern();
            scrollItemPattern.ScrollIntoView();
        }

        public static AutomationElement FindDescendentByConditionPath(this AutomationElement element, IEnumerable<Condition> conditionPath)
        {
            if (!conditionPath.Any())
            {
                return element;
            }

            var result = conditionPath.Aggregate(
                element,
                (parentElement, nextCondition) => parentElement == null
                                                      ? null
                                                      : parentElement.FindChildByCondition(nextCondition));

            return result;
        }

        public static AutomationElement FindDescendentById(this AutomationElement element, string id)
        {
            return element.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, id));
        }

        public static AutomationElement FindDescendentByIdPath(this AutomationElement element, IEnumerable<string> idPath)
        {
            var conditionPath = CreateConditionPathForPropertyValues(AutomationElement.AutomationIdProperty, idPath.Cast<object>());

            return FindDescendentByConditionPath(element, conditionPath);
        }

        public static AutomationElement FindDescendentByNamePath(this AutomationElement element, IEnumerable<string> namePath)
        {
            var conditionPath = CreateConditionPathForPropertyValues(AutomationElement.NameProperty, namePath.Cast<object>());

            return FindDescendentByConditionPath(element, conditionPath);
        }

        public static IEnumerable<Condition> CreateConditionPathForPropertyValues(AutomationProperty property, IEnumerable<object> values)
        {
            var conditions = values.Select(value => new PropertyCondition(property, value));

            return conditions.Cast<Condition>();
        }
        /// <summary>
        /// Finds the first child of the element that has a descendant matching the condition path.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <param name="conditionPath">The condition path.</param>
        /// <returns></returns>
        public static AutomationElement FindFirstChildHavingDescendantWhere(this AutomationElement element, IEnumerable<Condition> conditionPath)
        {
            return
                element.FindFirstChildHavingDescendantWhere(
                    child => child.FindDescendentByConditionPath(conditionPath) != null);
        }

        /// <summary>
        /// Finds the first child of the element that has a descendant matching the condition path.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <param name="conditionPath">The condition path.</param>
        /// <returns></returns>
        public static AutomationElement FindFirstChildHavingDescendantWhere(this AutomationElement element, Func<AutomationElement, bool> condition)
        {
            var children = element.FindAll(TreeScope.Children, Condition.TrueCondition);

            foreach (AutomationElement child in children)
            {
                if (condition(child))
                {
                    return child;
                }
            }

            return null;
        }

        public static AutomationElement FindChildById(this AutomationElement element, string automationId)
        {
            var result = element.FindChildByCondition(
                new PropertyCondition(AutomationElement.AutomationIdProperty, automationId));

            return result;
        }

        public static AutomationElement FindChildByName(this AutomationElement element, string name)
        {
            var result = element.FindChildByCondition(
                new PropertyCondition(AutomationElement.NameProperty, name));

            return result;
        }

        public static AutomationElement FindChildByClass(this AutomationElement element, string className)
        {
            var result = element.FindChildByCondition(
                new PropertyCondition(AutomationElement.ClassNameProperty, className));

            return result;
        }

        public static AutomationElement FindChildByProcessId(this AutomationElement element, int processId)
        {
            var result = element.FindChildByCondition(
                new PropertyCondition(AutomationElement.ProcessIdProperty, processId));

            return result;
        }

        public static AutomationElement FindChildByControlType(this AutomationElement element, ControlType controlType)
        {
            var result = element.FindChildByCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, controlType));

            return result;
        }

        public static AutomationElement FindChildByCondition(this AutomationElement element, Condition condition)
        {
            var result = element.FindFirst(
                TreeScope.Children,
                condition);

            return result;
        }

        /// <summary>
        /// Finds the child text block of an element.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <returns></returns>
        public static AutomationElement FindChildTextBlock(this AutomationElement element)
        {
            var child = TreeWalker.RawViewWalker.GetFirstChild(element);

            if (child != null && child.Current.ControlType == ControlType.Text)
            {
                return child;
            }

            return null;
        }
    }
}