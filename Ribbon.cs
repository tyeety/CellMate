using System;
using System.Drawing;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace CellMate
{
    [ComVisible(true)]
    public sealed class Ribbon : ExcelRibbon
    {
        private static readonly Bitmap Logo = Brand.Icon(32);
        public override string GetCustomUI(string ribbonId) => @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
 <ribbon><tabs><tab id='CellMateTab' label='CellMate'>
  <group id='CellMateTools' label='Tools'>
   <button id='PasteValues' label='Paste Values' size='large' imageMso='PasteValues' onAction='OnPasteValues' supertip='Paste only the values copied in Excel, keeping the destination formatting.'/>
   <button id='AutoFit' label='AutoFit Columns' size='large' imageMso='ColumnWidth' onAction='OnAutoFit'/>
  </group>
  <group id='CellMatePreferences' label='Preferences'>
   <button id='Settings' label='Settings' size='large' getImage='GetLogo' onAction='OnSettings' supertip='Customize or disable CellMate keyboard shortcuts.'/>
  </group>
 </tab></tabs></ribbon>
</customUI>";
        public object GetLogo(IRibbonControl control) => Logo;
        public void OnPasteValues(IRibbonControl control) => Queue("CellMate_PasteValues");
        public void OnAutoFit(IRibbonControl control) => Queue("CellMate_AutoFit");
        public void OnSettings(IRibbonControl control) => Queue("CellMate_Settings");
        private static void Queue(string command) => ExcelAsyncUtil.QueueAsMacro(() =>
        {
            // Run the registered command so Excel owns the macro/undo lifecycle.
            try { AddIn.Application.Run(command); }
            catch (Exception ex) { AddIn.ShowError(ex.Message); }
        });
    }
}
