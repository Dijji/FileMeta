//Copied from www.codeproject.com/Articles/101975/Building-a-Search-Text-Box-Control-with-WPF
//and modified to add next and previous buttons, and improve icons and general behaviour
//Licensed under The Code Project Open License (CPOL)
using System;
using System.Windows;
using System.Windows.Controls;

namespace FileMetadataAssociationManager
{

    public class ListBoxEx : ListBox
    {
        public enum ItemStyles{
            NormalStyle,
            CheckBoxStyle,
            RadioBoxStyle
        }
        private ItemStyles m_extendedStyle;

        public ItemStyles ExtendedStyle
        {
            get { return m_extendedStyle; }
            set {
                m_extendedStyle = value;

                // load resources
                ResourceDictionary resDict = new ResourceDictionary();
                resDict.Source = new Uri("pack://application:,,,/FileMetaAssociationManager;component/Themes/ListBoxEx.xaml");
                if (resDict.Source == null)
                    throw new SystemException();
 
                switch (value)
                {
                case ItemStyles.NormalStyle:
                    this.Style = (Style)resDict["NormalListBox"];
            	    break;
                case ItemStyles.CheckBoxStyle:
                    this.Style = (Style)resDict["CheckListBox"];
                    break; 
                case ItemStyles.RadioBoxStyle:
                    this.Style = (Style)resDict["RadioListBox"];
                    break;
                }
            }
        }

        static ListBoxEx()
        {
            DefaultStyleKeyProperty.OverrideMetadata(typeof(ListBoxEx), new FrameworkPropertyMetadata(typeof(ListBoxEx)));
        }

        public ListBoxEx(ItemStyles style)
            : base()
        {
            ExtendedStyle = style;
        }
        
        public ListBoxEx(): base(){
            ExtendedStyle = ItemStyles.NormalStyle;
        }


    }
}

