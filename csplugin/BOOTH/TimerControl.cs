using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BOOTH
{
    public class TimerControl : UserControl
    {
        protected SheetWriter writer;
        protected int number; 

        public TimerControl()
        {
            // NOTE: This constructor is here because the UI designer needs the base class
            // of a UI element to be non-abstract and to have a no-argument constructor.
            // This constructor should not be used in practice and this class should be treated
            // as abstract.
        }

        public TimerControl(SheetWriter writer, int number)
        {
            this.writer = writer;
            this.number = number;
        }

        public virtual string GetHeadingText()
        {
            throw new NotImplementedException();
        }

        public virtual void AddComment(string comment)
        {
            throw new NotImplementedException();
        }
    }
}
