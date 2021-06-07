using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH
{
    /**
     * MainPanelTimerControl is an abstract class for a timer that will be displayed on the
     * main timers panel. It is the assumption that timers placed there have start and stop buttons
     * and a note textbox (and the associated clear button) at least.
     */
    class MainPanelTimerControl : TimerControl
    {
        public MainPanelTimerControl()
        {
            // NOTE: This constructor is here because the UI designer needs the base class
            // of a UI element to be non-abstract and to have a no-argument constructor.
            // This constructor should not be used in practice and this class should be treated
            // as abstract.
        }

        public MainPanelTimerControl(SheetWriter writer, int number) : base(writer, number)
        {
        }
    }
}
