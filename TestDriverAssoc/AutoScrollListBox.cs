using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Controls;
using System.Windows;
using System.Collections.Specialized;

// Taken from http://blogs.microsoft.co.il/davidbi/2010/09/30/wpf-auto-scroll-listbox/
// No visible copyright

namespace TestDriverAssoc
{
    class AutoScrollListBox : ListBox
    {
        public bool AutoScroll
        {
            get { return (bool)GetValue(AutoScrollProperty); }
            set { SetValue(AutoScrollProperty, value); }
        }

        // Using a DependencyProperty as the backing store for AutoScoll.  This enables animation, styling, binding, etc...
        public static readonly DependencyProperty AutoScrollProperty =
            DependencyProperty.Register("AutoScroll", typeof(bool), typeof(AutoScrollListBox), new UIPropertyMetadata(default(bool), OnAutoScrollChanged));

        public static void OnAutoScrollChanged(DependencyObject s, DependencyPropertyChangedEventArgs e)
        {
            AutoScrollListBox thisLb = (AutoScrollListBox)s;

            // Add the event handler in case that the property is set to true
            if ((bool)e.NewValue == true && (bool)e.OldValue == false)
            {
                var ic = thisLb.Items as INotifyCollectionChanged;
                if (ic == null)
                {
                    return;
                }
                ic.CollectionChanged += new NotifyCollectionChangedEventHandler(thisLb.ic_CollectionChanged);
            }
            // Remove the event handel in case the property is set to false
            if ((bool)e.NewValue == false && (bool)e.OldValue == true)
            {
                var ic = thisLb.Items as INotifyCollectionChanged;
                if (ic == null)
                {
                    return;
                }
                ic.CollectionChanged -= new NotifyCollectionChangedEventHandler(thisLb.ic_CollectionChanged);
            }
        }

        void ic_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            var ic = sender as ItemCollection;
            if (ic != null)
            {
                //Scroll into the last item
                if (ic.Count > 1)
                {
                    this.ScrollIntoView(ic[ic.Count - 1]);
                }
            }
        }
    }
}
