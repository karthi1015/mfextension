
"""Reusable WPF forms for pyRevit."""

import sys
import os
import os.path as op
import string
from collections import OrderedDict
import threading
from functools import wraps

from pyrevit import HOST_APP, EXEC_PARAMS
from pyrevit.compat import safe_strtype
from pyrevit import coreutils
from pyrevit.coreutils.logger import get_logger
from pyrevit import framework
from pyrevit.framework import System
from pyrevit.framework import Threading
from pyrevit.framework import Interop
from pyrevit.framework import wpf, Forms, Controls, Media
from pyrevit.api import AdWindows
from pyrevit import revit, UI, DB


logger = get_logger(__name__)

DEFAULT_CMDSWITCHWND_WIDTH = 600
DEFAULT_SEARCHWND_WIDTH = 600
DEFAULT_SEARCHWND_HEIGHT = 100
DEFAULT_INPUTWINDOW_WIDTH = 500
DEFAULT_INPUTWINDOW_HEIGHT = 400


class WPFWindow(framework.Windows.Window):
    r"""WPF Window base class for all pyRevit forms.
    Args:
        xaml_source (str): xaml source filepath or xaml content
        literal_string (bool): xaml_source contains xaml content, not filepath
    Example:
        >>> from pyrevit import forms
        >>> layout = '<Window ShowInTaskbar="False" ResizeMode="NoResize" ' \
        >>>          'WindowStartupLocation="CenterScreen" ' \
        >>>          'HorizontalContentAlignment="Center">' \
        >>>          '</Window>'
        >>> w = forms.WPFWindow(layout, literal_string=True)
        >>> w.show()
    """

    def __init__(self, xaml_source, literal_string=False):
        """Initialize WPF window and resources."""
        # self.Parent = self
        wih = Interop.WindowInteropHelper(self)
        wih.Owner = AdWindows.ComponentManager.ApplicationWindow

        if not literal_string:
            if not op.exists(xaml_source):
                wpf.LoadComponent(self,
                                  os.path.join(EXEC_PARAMS.command_path,
                                               xaml_source)
                                  )
            else:
                wpf.LoadComponent(self, xaml_source)
        else:
            wpf.LoadComponent(self, framework.StringReader(xaml_source))

        #2c3e50 #noqa
        self.Resources['pyRevitDarkColor'] = \
            Media.Color.FromArgb(0xFF, 0x2c, 0x3e, 0x50)

        #23303d #noqa
        self.Resources['pyRevitDarkerDarkColor'] = \
            Media.Color.FromArgb(0xFF, 0x23, 0x30, 0x3d)

        #ffffff #noqa
        self.Resources['pyRevitButtonColor'] = \
            Media.Color.FromArgb(0xFF, 0xff, 0xff, 0xff)

        #f39c12 #noqa
        self.Resources['pyRevitAccentColor'] = \
            Media.Color.FromArgb(0xFF, 0xf3, 0x9c, 0x12)

        self.Resources['pyRevitDarkBrush'] = \
            Media.SolidColorBrush(self.Resources['pyRevitDarkColor'])
        self.Resources['pyRevitAccentBrush'] = \
            Media.SolidColorBrush(self.Resources['pyRevitAccentColor'])

        self.Resources['pyRevitDarkerDarkBrush'] = \
            Media.SolidColorBrush(self.Resources['pyRevitDarkerDarkColor'])

        self.Resources['pyRevitButtonForgroundBrush'] = \
            Media.SolidColorBrush(self.Resources['pyRevitButtonColor'])

    def show(self, modal=False):
        """Show window."""
        if modal:
            return self.ShowDialog()
        self.Show()

    def show_dialog(self):
        """Show modal window."""
        return self.ShowDialog()

    def set_image_source(self, element_name, image_file):
        """Set source file for image element.
        Args:
            element_name (str): xaml image element name
            image_file (str): image file path
        """
        wpfel = getattr(self, element_name)
        if not op.exists(image_file):
            # noinspection PyUnresolvedReferences
            wpfel.Source = \
                framework.Imaging.BitmapImage(
                    framework.Uri(os.path.join(EXEC_PARAMS.command_path,
                                               image_file))
                    )
        else:
            wpfel.Source = \
                framework.Imaging.BitmapImage(framework.Uri(image_file))

    @staticmethod
    def hide_element(*wpf_elements):
        """Collapse elements.
        Args:
            *wpf_elements (str): element names to be collaped
        """
        for wpfel in wpf_elements:
            wpfel.Visibility = framework.Windows.Visibility.Collapsed

    @staticmethod
    def show_element(*wpf_elements):
        """Show collapsed elements.
        Args:
            *wpf_elements (str): element names to be set to visible.
        """
        for wpfel in wpf_elements:
            wpfel.Visibility = framework.Windows.Visibility.Visible

    @staticmethod
    def toggle_element(*wpf_elements):
        """Toggle visibility of elements.
        Args:
            *wpf_elements (str): element names to be toggled.
        """
        for wpfel in wpf_elements:
            if wpfel.Visibility == framework.Windows.Visibility.Visible:
                self.hide_element(wpfel)
            elif wpfel.Visibility == framework.Windows.Visibility.Collapsed:
                self.show_element(wpfel)


class TemplateUserInputWindow(WPFWindow):
    """Base class for pyRevit user input standard forms.
    Args:
        context (any): window context element(s)
        title (str): window title
        width (int): window width
        height (int): window height
        **kwargs: other arguments to be passed to :func:`_setup`
    """

    xaml_source = 'BaseWindow.xaml'

    def __init__(self, context, title, width, height, **kwargs):
        """Initialize user input window."""
        WPFWindow.__init__(self,
                           op.join(op.dirname(__file__), self.xaml_source))
        self.Title = title
        self.Width = width
        self.Height = height

        self._context = context
        self.response = None
        self.PreviewKeyDown += self.handle_input_key

        self._setup(**kwargs)

    def _setup(self, **kwargs):
        """Private method to be overriden by subclasses for window setup."""
        pass

    def handle_input_key(self, sender, args):
        """Handle keyboard input."""
        if args.Key == framework.Windows.Input.Key.Escape:
            self.Close()

    @classmethod
    def show(cls, context,
             title='User Input',
             width=DEFAULT_INPUTWINDOW_WIDTH,
             height=DEFAULT_INPUTWINDOW_HEIGHT, **kwargs):
        """Show user input window.
        Args:
            context (any): window context element(s)
            title (type): window title
            width (type): window width
            height (type): window height
            **kwargs (type): other arguments to be passed to window
        """
        dlg = cls(context, title, width, height, **kwargs)
        dlg.ShowDialog()
        return dlg.response


class SelectFromList(TemplateUserInputWindow):
    """Standard form to select from a list of items.
    Args:
        context (list[str]): list of items to be selected from
        title (str): window title
        width (int): window width
        height (int): window height
        button_name (str): name of select button
        multiselect (bool): allow multi-selection
    Example:
        >>> from pyrevit import forms
        >>> items = ['item1', 'item2', 'item3']
        >>> forms.SelectFromList.show(items, button_name='Select Item')
        >>> ['item1']
    """

    xaml_source = 'SelectFromList.xaml'

    def _setup(self, **kwargs):
        self.hide_element(self.clrsearch_b)
        self.clear_search(None, None)
        self.search_tb.Focus()

        if 'multiselect' in kwargs and not kwargs['multiselect']:
            self.list_lb.SelectionMode = Controls.SelectionMode.Single
        else:
            self.list_lb.SelectionMode = Controls.SelectionMode.Extended

        button_name = kwargs.get('button_name', None)
        if button_name:
            self.select_b.Content = button_name

        self._list_options()

    def _list_options(self, option_filter=None):
        if option_filter:
            option_filter = option_filter.lower()
            self.list_lb.ItemsSource = \
                [safe_strtype(option) for option in self._context
                 if option_filter in safe_strtype(option).lower()]
        else:
            self.list_lb.ItemsSource = \
                [safe_strtype(option) for option in self._context]

    def _get_options(self):
        return [option for option in self._context
                if safe_strtype(option) in self.list_lb.SelectedItems]

    def button_select(self, sender, args):
        """Handle select button click."""
        self.response = self._get_options()
        self.Close()

    def search_txt_changed(self, sender, args):
        """Handle text change in search box."""
        if self.search_tb.Text == '':
            self.hide_element(self.clrsearch_b)
        else:
            self.show_element(self.clrsearch_b)

        self._list_options(option_filter=self.search_tb.Text)

    def clear_search(self, sender, args):
        """Clear search box."""
        self.search_tb.Text = ' '
        self.search_tb.Clear()
        self.search_tb.Focus()


class BaseCheckBoxItem(object):
    """Base class for checkbox option wrapping another object."""

    def __init__(self, orig_item):
        """Initialize the checkbox option and wrap given obj.
        Args:
            orig_item (any): object to wrap (must have name property
                             or be convertable to string with str()
        """
        self.item = orig_item
        self.state = False

    def __nonzero__(self):
        return self.state

    def __str__(self):
        return self.name or str(self.item)

    @property
    def name(self):
        """Name property."""
        return getattr(self.item, 'name', '')

    def unwrap(self):
        """Unwrap and return wrapped object."""
        return self.item

		
class SelectFromDoubleList(TemplateUserInputWindow):
    """Standard form to select from a list of items.
    Args:
        context (list[str]): list of items to be selected from
        title (str): window title
        width (int): window width
        height (int): window height
        button_name (str): name of select button
        multiselect (bool): allow multi-selection
    Example:
        >>> from pyrevit import forms
        >>> items = ['item1', 'item2', 'item3']
        >>> forms.SelectFromList.show(items, button_name='Select Item')
        >>> ['item1']
    """

    xaml_source = 'SelectFromDoubleList.xaml'

    def _setup(self, **kwargs):
        self.hide_element(self.clrsearch_b)
        self.clear_search(None, None)
        self.search_tb.Focus()

        if 'multiselect' in kwargs and not kwargs['multiselect']:
            self.list_lb.SelectionMode = Controls.SelectionMode.Single
        else:
            self.list_lb.SelectionMode = Controls.SelectionMode.Extended

        button_name = kwargs.get('button_name', None)
        if button_name:
            self.select_b.Content = button_name

        self._list_options()

    def _list_options(self, option_filter=None):
        if option_filter:
            option_filter = option_filter.lower()
            self.list_lb.ItemsSource = \
                [safe_strtype(option) for option in self._context
                 if option_filter in safe_strtype(option).lower()]
        else:
            self.list_lb.ItemsSource = \
                [safe_strtype(option) for option in self._context]

    def _get_options(self):
        return [option for option in self._context
                if safe_strtype(option) in self.list_lb.SelectedItems]

    def button_select(self, sender, args):
        """Handle select button click."""
        self.response = self._get_options()
        self.Close()

    def search_txt_changed(self, sender, args):
        """Handle text change in search box."""
        if self.search_tb.Text == '':
            self.hide_element(self.clrsearch_b)
        else:
            self.show_element(self.clrsearch_b)

        self._list_options(option_filter=self.search_tb.Text)

    def clear_search(self, sender, args):
        """Clear search box."""
        self.search_tb.Text = ' '
        self.search_tb.Clear()
        self.search_tb.Focus()

		