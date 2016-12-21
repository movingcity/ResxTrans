using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;

namespace ResxTrans
{
    public static class ViewHelper
    {
        public static async Task<MessageDialogResult> ShowMessageAsync(string title, string message, MetroWindow window = null)
        {
            MetroWindow metroWindow;
            if (window == null)
            {
                metroWindow = Application.Current.MainWindow as MetroWindow;
            }
            else
            {
                metroWindow = window;
            }
            var result = await metroWindow.ShowMessageAsync(title, message);
            return result;
        }

        public static async Task<ProgressDialogController> ShowProgressAsync(string title, string message, MetroWindow window = null)
        {
            MetroWindow metroWindow;
            if (window == null)
            {
                metroWindow = Application.Current.MainWindow as MetroWindow;
            }
            else
            {
                metroWindow = window;
            }
            var result = await metroWindow.ShowProgressAsync(title, message);
            return result;
        }

        public static async Task<MessageDialogResult> ShowYesNoDialog(string title, string message, MetroWindow window = null)
        {
            MetroWindow metroWindow;
            if (window == null)
            {
                metroWindow = Application.Current.MainWindow as MetroWindow;
            }
            else
            {
                metroWindow = window;
            }
            var mySettings = new MetroDialogSettings
            {
                AffirmativeButtonText = "是",
                NegativeButtonText = "否"
            };
            var result = await metroWindow.ShowMessageAsync(title, message, MessageDialogStyle.AffirmativeAndNegative, mySettings);
            return result;
        }
    }
}
