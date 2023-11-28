using sapfewse;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.SAP
{
    public class SapComWrapper<T>
    {
        public T Com { get; private set; }
        public string Text
        {
            get => ((GuiVComponent)Com).Text;
            set => ((GuiVComponent)Com).Text = value;
        }

        //public string Key
        //{
        //    get => (Com as ISapComboBoxTarget).Key;
        //    set => (Com as ISapComboBoxTarget).Key = value;
        //}

        //public int CaretPosition
        //{
        //    get => (Com as ISapCTextField).CaretPosition;
        //    set => (Com as ISapCTextField).CaretPosition = value;
        //}

        public SapComWrapper(T nativeComObejct)
        {
            Com = nativeComObejct;
        }

        public SapComWrapper<U> FindById<U>(string pathId, bool showTypes = false)
        {
            return Sap.FindById<U>((GuiComponent)Com, pathId, showTypes);
        }

        //public U To<U>()
        //{
        //    return (U)Com;
        //}

        //public void Iconify() => (Com as GuiFrameWindow).Iconify();

        public virtual void SendVKey(int vKey) => ((ISapWindowTarget)Com).SendVKey(vKey);

        public void Close() => ((ISapWindowTarget)Com).Close();

        public void Maximize() => ((ISapWindowTarget)Com).Maximize();

        public void Press() => ((ISapButtonTarget)Com).Press();

        public void SetFocus() => ((GuiVComponent)Com).SetFocus();

        //public void Select() => (Com as ISapTabTarget).Select();

        //public void SelectAll() => (Com as ISapGridViewTarget).SelectAll();
    }
}
