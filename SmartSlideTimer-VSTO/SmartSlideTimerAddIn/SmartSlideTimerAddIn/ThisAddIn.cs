using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace SmartSlideTimerAddIn
{
    public partial class ThisAddIn
    {
        private float GetSlideTiming(PowerPoint.Shapes shapes)
        {
            float wordcount = 0;

            foreach (PowerPoint.Shape shape in shapes)
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    wordcount += shape.TextFrame.TextRange.Text.Split().Count();
                }
            }

            return wordcount;
        }

        private void SetSlideShowAdvanceTime(PowerPoint.Presentation presentation)
        {
            foreach (PowerPoint.Slide Sld in presentation.Application.ActivePresentation.Slides)
            {
                Sld.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoTrue;
                Sld.SlideShowTransition.AdvanceTime = this.GetSlideTiming(Sld.Shapes);
            }

            presentation.Save();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.PresentationSave += new PowerPoint.EApplication_PresentationSaveEventHandler(this.SetSlideShowAdvanceTime);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
