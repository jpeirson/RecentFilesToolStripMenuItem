#region

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;

#endregion

namespace RecentFilesMenuItem
{
    /// <summary>
    ///     Represents a menu for recently-access files to be displayed within a System.Windows.Forms.MainMenu.
    /// </summary>
    internal class RecentFilesMenuItem : MenuItem
    {
        private readonly MenuItem _clearMenuItem = new MenuItem();
        private readonly List<RecentMenuItem> _items = new List<RecentMenuItem>();
        private readonly MenuItem _openAllMenuItem = new MenuItem();

        private string _clearOptionText = "Clear All Recent Items";
        private bool _displayClearOption = true;
        private RecentFilesDisplayMode _displayMode = RecentFilesDisplayMode.Child;
        private int _maxDisplayItems = 10;
        private string _openAllOptionText = "Open All Recent Items";
        private bool _prependItemNumbers = true;

        /// <summary>
        ///     Initializes a new RecentFilesMenuItem.
        /// </summary>
        public RecentFilesMenuItem()
        {
            _openAllMenuItem.Text = OpenAllOptionText;
            _openAllMenuItem.Click += delegate
            {
                if (AllItemClicked != null)
                    AllItemClicked(this, EventArgs.Empty);
            };

            _clearMenuItem.Text = ClearOptionText;
            _clearMenuItem.Click += delegate
            {
                Clear();

                if (ClearItemClicked != null)
                    ClearItemClicked(this, EventArgs.Empty);
            };
        }

        /// <summary>
        ///     Returns a list of all current items.
        /// </summary>
        public ReadOnlyCollection<RecentMenuItem> Items
        {
            get { return _items.AsReadOnly(); }
        }

        /// <summary>
        ///     Gets or sets whether the 'clear all' option should be displayed.
        /// </summary>
        [Description("Indicates whether the 'clear all' option should be displayed."), Category("Appearance")]
        public bool DisplayClearOption
        {
            get { return _displayClearOption; }
            set
            {
                _displayClearOption = value;
                PopulateItems();
            }
        }

        /// <summary>
        ///     Gets or sets the text to display for 'clear all' option.
        /// </summary>
        [Description("Text to display for 'clear all' option."), Category("Data")]
        public string ClearOptionText
        {
            get { return _clearOptionText; }
            set
            {
                _clearOptionText = value;
                _clearMenuItem.Text = value;
            }
        }

        /// <summary>
        ///     Gets or sets whether the 'open all' option should be displayed.
        /// </summary>
        [Description("Indicates whether the 'open all' option should be displayed."), Category("Appearance")]
        public bool DisplayOpenAllOption
        {
            get { return _displayOpenAllOption; }
            set
            {
                _displayOpenAllOption = value;
                PopulateItems();
            }
        }

        /// <summary>
        ///     Gets or sets the text to display for 'open all' option.
        /// </summary>
        [Description("Text to display for 'open all' option."), Category("Data")]
        public string OpenAllOptionText
        {
            get { return _openAllOptionText; }
            set
            {
                _openAllOptionText = value;
                _openAllMenuItem.Text = value;
            }
        }

        /// <summary>
        ///     Gets or sets the current display mode.
        /// </summary>
        [Description("Indicates which display mode to use."), Category("Behavior")]
        public RecentFilesDisplayMode DisplayMode
        {
            get { return _displayMode; }
            set
            {
                _displayMode = value;
                PopulateItems();
            }
        }

        /// <summary>
        ///     Gets or sets the maximum number of items to display.
        /// </summary>
        [Description("Indicates the maximum number of items to display."), Category("Behavior")]
        public int MaxDisplayItems
        {
            get { return _maxDisplayItems; }
            set
            {
                _maxDisplayItems = value;

                if (_items.Count > _maxDisplayItems)
                    _items.RemoveRange(_maxDisplayItems, _items.Count - _maxDisplayItems);

                PopulateItems();
            }
        }

        /// <summary>
        ///     Gets or sets whether to prepend human-friendly number indicatators to menu items.
        /// </summary>
        [Description("Indicates whether to prepend human-friendly number indicatators to menu items."), Category("Behavior")]
        public bool PrependItemNumbers
        {
            get { return _prependItemNumbers; }
            set
            {
                _prependItemNumbers = value;
                PopulateItems();
            }
        }

        /// <summary>
        ///     Occurs when the menu item is clicked or selected using a shortcut key or access key defined for the menu item.
        /// </summary>
        public event EventHandler ItemClicked;

        /// <summary>
        ///     Occurs when the 'clear all' menu item is clicked or selected using a shortcut key or access key defined for the
        ///     menu item.
        /// </summary>
        public event EventHandler ClearItemClicked;

        /// <summary>
        ///     Occurs when the 'open all' menu item is clicked or selected using a shortcut key or access key defined for the menu
        ///     item.
        /// </summary>
        public event EventHandler AllItemClicked;

        /// <summary>
        ///     Adds a new menu item to the menu.
        /// </summary>
        /// <param name="menuItem">The menu item to add.</param>
        public void Add(RecentMenuItem menuItem)
        {
            if (!menuItem.FileInfo.Exists)
                return;

            Remove(menuItem, false);

            if (_items.Count == _maxDisplayItems)
                _items.RemoveAt(_items.Count - 1);

            _items.Insert(0, menuItem);

            menuItem.Click += (s, e) =>
            {
                if (ItemClicked != null)
                    ItemClicked(s, e);
            };

            PopulateItems();
        }

        /// <summary>
        ///     Removes a menu item from the menu
        /// </summary>
        /// <param name="menuItem">The menu item to remove.</param>
        /// <param name="removeExistingFiles">Indicates whether menu items with the same file path should be removed as well.</param>
        /// <param name="repopulateDisplayItems">Indicates whether the menu should be repopulated afterwards.</param>
        public void Remove(RecentMenuItem menuItem, bool removeExistingFiles, bool repopulateDisplayItems = true)
        {
            _items.Remove(menuItem);

            if (removeExistingFiles)
                _items.RemoveAll(x => x.FileInfo.FullName.Equals(menuItem.FileInfo.FullName, StringComparison.OrdinalIgnoreCase));

            if (repopulateDisplayItems)
                PopulateItems();
        }

        /// <summary>
        ///     Clears all items in the menu.
        /// </summary>
        public void Clear()
        {
            _items.Clear();

            if (DisplayMode == RecentFilesDisplayMode.Child)
                ClearChildItems();

            if (DisplayMode == RecentFilesDisplayMode.Consecutive)
                ClearConsecutiveItems();
        }

        private void PopulateItems()
        {
            // update item numbers
            for (var i = 0; i < _items.Count; i++)
            {
                var item = _items[i];
                var itemText = string.IsNullOrEmpty(item.DisplayText) ? item.FileInfo.FullName : item.DisplayText;
                item.Text = PrependItemNumbers ? string.Format("{0}: {1}", i + 1, itemText) : itemText;
            }

            ClearChildItems();
            ClearConsecutiveItems();

            if (DisplayMode == RecentFilesDisplayMode.Child)
                PopulateChildItems();

            if (DisplayMode == RecentFilesDisplayMode.Consecutive)
                PopulateConsecutiveItems();
        }

        #region RecentFilesDisplayMode.Child Methods

        private void PopulateChildItems()
        {
            if (_items.Count == 0)
                return;

            for (var i = 0; i < _items.Count; i++)
            {
                if (i == MaxDisplayItems)
                    break;

                PopulateChildItem(_items[i]);
            }

            if (DisplayOpenAllOption || DisplayClearOption)
                MenuItems.Add(new MenuItemSeperator());

            if (DisplayOpenAllOption)
                MenuItems.Add(_openAllMenuItem);

            if (DisplayClearOption)
                MenuItems.Add(_clearMenuItem);
        }

        private void PopulateChildItem(MenuItem item)
        {
            Visible = true;
            MenuItems.Add(0, item);
            Enabled = MenuItems.Count > 0;
        }

        private void ClearChildItems()
        {
            MenuItems.Clear();
            Enabled = false;
        }

        #endregion

        #region RecentFilesDisplayMode.Consecutive Methods

        private readonly List<int> _consecutiveItemIndexes = new List<int>();
        private bool _displayOpenAllOption;

        private void PopulateConsecutiveItems()
        {
            Visible = false;

            if (_items.Count == 0)
                return;

            var itemIndex = Parent.MenuItems.IndexOf(this);

            PopulateConsecutiveItem(itemIndex, new MenuItemSeperator());
            itemIndex++;

            foreach (var item in _items)
            {
                if (itemIndex == MaxDisplayItems)
                    break;

                PopulateConsecutiveItem(itemIndex, item);
                itemIndex++;
            }

            if (DisplayOpenAllOption || DisplayClearOption)
            {
                PopulateConsecutiveItem(itemIndex, new MenuItemSeperator());
                itemIndex++;
            }

            if (DisplayOpenAllOption)
            {
                PopulateConsecutiveItem(itemIndex, _openAllMenuItem);
                itemIndex++;
            }

            if (DisplayClearOption)
            {
                PopulateConsecutiveItem(itemIndex, _clearMenuItem);
                itemIndex++;
            }

            PopulateConsecutiveItem(itemIndex, new MenuItemSeperator());
        }

        private void PopulateConsecutiveItem(int index, MenuItem item)
        {
            Parent.MenuItems.Add(index, item);
            _consecutiveItemIndexes.Add(index);
        }

        private void ClearConsecutiveItems()
        {
            var owner = (Parent as MenuItem);
            _consecutiveItemIndexes.Reverse();

            foreach (var index in _consecutiveItemIndexes)
            {
                owner.MenuItems.RemoveAt(index);
            }

            _consecutiveItemIndexes.Clear();
        }

        #endregion

        /// <summary>
        ///     Represents an individual menu item separator.
        /// </summary>
        internal class MenuItemSeperator : MenuItem
        {
            public MenuItemSeperator()
            {
                Text = "-";
            }
        }

        /// <summary>
        ///     Represents an individual recent menu item.
        /// </summary>
        public class RecentMenuItem : MenuItem
        {
            /// <summary>
            ///     Initializes a new RecentMenuItem.
            /// </summary>
            /// <param name="fileInfo">The respective FileInfo for the menu item.</param>
            public RecentMenuItem(FileInfo fileInfo)
            {
                FileInfo = fileInfo;
            }

            /// <summary>
            ///     Respective file info reference.
            /// </summary>
            public FileInfo FileInfo { get; private set; }

            /// <summary>
            ///     Text to display on menu item.
            /// </summary>
            public string DisplayText { get; set; }
        }

        #region RecentFilesDisplayMode enum

        /// <summary>
        ///     Display mode for menu item.
        /// </summary>
        public enum RecentFilesDisplayMode
        {
            /// <summary>
            ///     Items are displayed using a parent/child model.
            /// </summary>
            Child,

            /// <summary>
            ///     Iteams are displayed consecutively, being appended after the parent item.
            /// </summary>
            Consecutive
        }

        #endregion
    }
}