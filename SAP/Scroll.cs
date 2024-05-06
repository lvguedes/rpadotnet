using sapfewse;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RpaLib.Tracing;
using System.Threading;

namespace RpaLib.SAP
{
    public class Scroll
    {
        const int DELAY_AFTER_SCROLL = 0;
        public GuiScrollbar GuiScrollbar { get; private set; }

        public int Position
        {
            get => GuiScrollbar.Position;
            set => GuiScrollbar.Position = value;
        }

        public int Minimum { get => GuiScrollbar.Minimum; }

        public int Maximum { get => GuiScrollbar.Maximum; }

        public int PageSize { get => GuiScrollbar.PageSize; }

        public Scroll (GuiScrollbar guiScrollbar, int delayAfterScroll = DELAY_AFTER_SCROLL)
        {
            GuiScrollbar = guiScrollbar;
        }

        public bool IsNeeded() => IsNeeded(GuiScrollbar);
        public bool IsNeeded(int currentRow) => IsNeeded(GuiScrollbar, currentRow);
        public void NextPage(int pages = 1, int delayAfterScroll = DELAY_AFTER_SCROLL) => NextPage(GuiScrollbar);
        public void Reset() => Reset(GuiScrollbar);
        public void ShowCurrentStateInfo() => ShowCurrentStateInfo(GuiScrollbar);

        public static bool IsNeeded(GuiScrollbar scrollbar)
        {
            if (scrollbar.Position < scrollbar.Maximum)
                return true;
            else
                return false;
        }

        public static bool IsNeeded(GuiScrollbar scrollbar, int currentRow)
        {
            if (IsNeeded(scrollbar) && currentRow != 0 && currentRow % scrollbar.PageSize == 0)
                return true;
            else
                return false;
        }

        public static void NextPage(GuiScrollbar scrollbar, int pages = 1, int delayAfterScroll = DELAY_AFTER_SCROLL)
        {
            scrollbar.Position += scrollbar.PageSize * pages;
            Thread.Sleep(delayAfterScroll);
        }

        public static void Reset(GuiScrollbar scrollbar)
        {
            scrollbar.Position = 0;
        }

        public static void ShowCurrentStateInfo(GuiScrollbar scrollbar)
        {
            Trace.WriteLine(
                string.Join("\n",
                    $"Scrollbar properties:",
                    $"  Position: {scrollbar.Position}", // index of the current state (leftmost column or topmost row showing)
                    $"  Minimum:  {scrollbar.Minimum}",  // index of first filled line
                    $"  Maximum:  {scrollbar.Maximum}",  // index of last filled line
                    $"  PageSize: {scrollbar.PageSize}", // number of lines (filled or not) supported by each page
                    $"  Seq:    {scrollbar.Range}"     // same as Maximum
                )
            );
        }
    }
}
