namespace PDFReporter
{
    using iTextSharp.text;

    public static class PDFStyle
    {
        public static Font SetFontStyle(string fontFamily, int fontSize, int fontWeigth)
        {
            return FontFactory.GetFont(fontFamily, fontSize, fontWeigth);
        }

        public static BaseColor Color(int red, int green, int blue)
        {
            return new BaseColor(red, green, blue);
        }
    }
}