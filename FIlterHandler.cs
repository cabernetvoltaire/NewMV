public class FilterHandler
{

    public enum FilterState : byte
    {

        All,

        PicVid,

        Piconly,

        Vidonly,

        LinkOnly,

        NoPicVid,
    }

    private List<string> mFileList;

    public List<string> FileList
    {
        get
        {
            return FilterLBList();
        }
        set
        {
            mFileList = value;
        }
    }

    public event EventHandler StateChanged;

    private object sender;

    private EventArgs e;

    public Color Colour
    {
        get
        {
            return mColour(mState);
        }
    }

    private List<string> mDescList = new List<string>();

    public List<string> Descriptions
    {
        get
        {
            for (i = 0; (i <= 5); i++)
            {
                mDescList.Add(mDescription(i));
            }

            return mDescList;
        }
    }

    public string Description
    {
        get
        {
            return mDescription(mState);
        }
    }

    private byte mState;

    public FilterHandler()
    {
        this.State = FilterState.All;
    }

    public byte OldState
    {
    }

    private byte value;
}
EndPropertyEndclass Unknown
{
}


public void IncrementState()
{
    this.State = ((this.State + 1)
                % (FilterState.NoPicVid + 1));
}

private List<string> FilterLBList()
{
    List<string> lst = new List<string>();
    foreach (string f in mFileList)
    {
        lst.Add(f);
    }

    switch (CurrentfilterState.State)
    {
        case FilterHandler.FilterState.All:
            break;
        case FilterHandler.FilterState.NoPicVid:
            foreach (string m in mFileList)
            {
                IO.FileInfo f = new IO.FileInfo(m);
                if (((((PICEXTENSIONS + VIDEOEXTENSIONS).IndexOf(f.Extension.ToLower()) + 1)
                            == 0)
                            && (f.Extension.Length != 0)))
                {

                }
                else
                {
                    lst.Remove(m);
                }

            }

            break;
        case FilterHandler.FilterState.LinkOnly:
            foreach (string m in mFileList)
            {
                IO.FileInfo f = new IO.FileInfo(m);
                if ((f.Extension.ToLower() == ".lnk"))
                {

                }
                else
                {
                    lst.Remove(m);
                }

            }

            break;
        case FilterHandler.FilterState.Piconly:
            foreach (string m in mFileList)
            {
                IO.FileInfo f = new IO.FileInfo(m);
                if ((((PICEXTENSIONS.IndexOf(f.Extension.ToLower()) + 1)
                            == 0)
                            && (f.Extension.Length != 0)))
                {
                    lst.Remove(m);
                }
                else
                {

                }

            }

            break;
        case FilterHandler.FilterState.PicVid:
            foreach (string m in mFileList)
            {
                IO.FileInfo f = new IO.FileInfo(m);
                if (((((PICEXTENSIONS + VIDEOEXTENSIONS).IndexOf(f.Extension.ToLower()) + 1)
                            == 0)
                            && (f.Extension.Length != 0)))
                {
                    lst.Remove(m);
                }
                else
                {

                }

            }

            break;
        case FilterHandler.FilterState.Vidonly:
            foreach (string m in mFileList)
            {
                IO.FileInfo f = new IO.FileInfo(m);
                if ((((VIDEOEXTENSIONS.IndexOf(f.Extension.ToLower()) + 1)
                            == 0)
                            && (f.Extension.Length != 0)))
                {
                    lst.Remove(m);
                }
                else
                {

                }

            }

            break;
    }
    return lst;
}