using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TDV.Docx
{
    public enum LINE_TYPE
    {
        NONE,
        SINGLE,
        DOTTED

        /*
        single - a single line
dashDotStroked - a line with a series of alternating thin and thick strokes
dashed - a dashed line
dashSmallGap - a dashed line with small gaps
dotDash - a line with alternating dots and dashes
dotDotDash - a line with a repeating dot - dot - dash sequence
dotted - a dotted line
double - a double line
doubleWave - a double wavy line
inset - an inset set of lines
nil - no border
none - no border
outset - an outset set of lines
thick - a single line
thickThinLargeGap - a thick line contained within a thin line with a large-sized intermediate gap
thickThinMediumGap - a thick line contained within a thin line with a medium-sized intermediate gap
thickThinSmallGap - a thick line contained within a thin line with a small intermediate gap
thinThickLargeGap - a thin line contained within a thick line with a large-sized intermediate gap
thinThickMediumGap - a thick line contained within a thin line with a medium-sized intermediate gap
thinThickSmallGap - a thick line contained within a thin line with a small intermediate gap
thinThickThinLargeGap - a thin-thick-thin line with a large gap
thinThickThinMediumGap - a thin-thick-thin line with a medium gap
thinThickThinSmallGap - a thin-thick-thin line with a small gap
threeDEmboss - a three-staged gradient line, getting darker towards the paragraph
threeDEngrave - a three-staged gradient like, getting darker away from the paragraph
triple - a triple line
wave - a wavy line
         */
    }

    public static class EnumExtentions
    {
        public static string ToStringValue(this LINE_TYPE e)
        {
            switch (e)
            {
                case LINE_TYPE.DOTTED:
                    return "dotted";
                case LINE_TYPE.SINGLE:
                    return "single";
                case LINE_TYPE.NONE:
                    return "";
            }
            throw new NotImplementedException();
        }
        public static string ToStringValue(this XML_SPACE e)
        {
            switch (e)
            {
                case XML_SPACE.DEFAULT:
                    return "default";
                case XML_SPACE.PRESERVE:
                    return "preserve";
            }
            throw new NotImplementedException();
        }
        public static string ToStringValue(this PAGE_ORIENTATION e)
        {
            switch (e)
            {
                case PAGE_ORIENTATION.PORTRAIT:
                    return "portrait";
                case PAGE_ORIENTATION.LANSCAPE:
                    return "landscape";
            }
            throw new NotImplementedException();
        }
        public static string ToStringValue(this NUM_FMT e)
        {
            switch (e)
            {
                case NUM_FMT.NONE:
                    return "none";
                case NUM_FMT.DEFAULT:
                    return "";
                case NUM_FMT.UPPER_LETTER:
                    return "upperLetter";
                case NUM_FMT.LOWER_ROMAN:
                    return "lowerRoman";
                case NUM_FMT.UPPER_ROMAN:
                    return "upperRoman";
                case NUM_FMT.CHICAGO:
                    return "chicago";
                case NUM_FMT.BULLET:
                    return "bullet";
                case NUM_FMT.CARDINAL_TEXT:
                    return "cardinalText";
                case NUM_FMT.DECIMAL:
                    return "decimal";
                case NUM_FMT.DECIMAL_ENCLOSED_CIRCLE:
                    return "decimalEnclosedCircle";
                case NUM_FMT.DECIMAL_ENCLOSED_FULL_STOP:
                    return "decimalEnclosedFullstop ";
                case NUM_FMT.DECIMAL_ENCLOSED_PAREN:
                    return "decimalEnclosedParen";
                case NUM_FMT.DECIMAL_ZERO:
                    return "decimalZero";
                case NUM_FMT.ORDINAL_TEXT:
                    return "ordinalText";
                case NUM_FMT.IDEOGRAPN_DIGITAL:
                    return "ideographDigital";
            }
            throw new NotImplementedException();
        }
        public static string ToStringValue(this FLD_CHAR_TYPE e)
        {
            switch (e)
            {
                case FLD_CHAR_TYPE.BEGIN:
                    return "begin";
                case FLD_CHAR_TYPE.SEPARATE:
                    return "separate";
                case FLD_CHAR_TYPE.END:
                    return "end";
            }
            throw new NotImplementedException();
        }
        public static string ToStringValue(this FOOTER_TYPE e)
        {
            switch (e)
            {
                case FOOTER_TYPE.NONE:
                    return "";
                case FOOTER_TYPE.SEPARATOR:
                    return "separator";
                case FOOTER_TYPE.CONTINUATION_SEPAPRATOR:
                    return "continuationSeparator";
            }
            throw new NotImplementedException();
        }
        public static string ToStringValue(this DOC_PART_GALLERY_VALUE e)
        {
            switch (e)
            {
                case DOC_PART_GALLERY_VALUE.NONE:
                    return "";
                case DOC_PART_GALLERY_VALUE.PAGE_NUMBERS_TOP_OF_PAGE:
                    return "Page Numbers (Top of Page)";
                case DOC_PART_GALLERY_VALUE.PAGE_NUMBERS_BOTTOM_OF_PAGE:
                    return "Page Numbers (Bottom of Page)";
            }
            throw new NotImplementedException();
        }
        public static string ToStringValue(this REFERENCE_TYPE e)
        {
            switch (e)
            {
                case REFERENCE_TYPE.FIRST:
                    return "first";
                case REFERENCE_TYPE.EVEN:
                    return "even";
                case REFERENCE_TYPE.DEFAULT:
                    return "default";
            }
            throw new NotImplementedException();
        }
        public static string ToStringValue(this HORIZONTAL_ALIGN e)
        {
            switch (e)
            {
                case HORIZONTAL_ALIGN.LEFT:
                    return "left";
                case HORIZONTAL_ALIGN.CENTER:
                    return "center";
                case HORIZONTAL_ALIGN.RIGHT:
                    return "right";
                case HORIZONTAL_ALIGN.BOTH:
                    return "both";
            }
            throw new NotImplementedException();
        }
        public static string ToStringValue(this MULTI_LEVEL_TYPE e)
        {
            switch (e)
            {
                case MULTI_LEVEL_TYPE.SINGLE_LEVEL:
                    return "singleLevel";
                case MULTI_LEVEL_TYPE.MULTI_LEVEL:
                    return "multiLevel";
                case MULTI_LEVEL_TYPE.HYBRID_MULTY_LEVEL:
                    return "hybridMultiLevel";
            }
            throw new NotImplementedException();
        }

        public static string ToStringValue(this BORDER e)
        {
            switch (e)
            {
                case BORDER.LEFT:
                    return "left";
                case BORDER.RIGHT:
                    return "right";
                case BORDER.TOP:
                    return "top";
                case BORDER.BOTTOM:
                    return "bottom";
                case BORDER.BETWEEN:
                    return "between";
                case BORDER.BAR:
                    return "bar";
                case BORDER.INSIDE_H:
                    return "insideH";
                case BORDER.INSIDE_V:
                    return "insideV";
            }
            throw new NotImplementedException();
        }
        public static T ToEnum<T>(string value)
        {
            Type currType = typeof(T);
            if (currType == typeof(LINE_TYPE))
            {
                switch (value)
                {
                    case "single":
                        return (T)(object)LINE_TYPE.SINGLE;
                    case "dotted":
                        return (T)(object)LINE_TYPE.DOTTED;
                    case "nil":
                    case "":
                        return (T)(object)LINE_TYPE.NONE;
                }
            }
            else if (currType == typeof(XML_SPACE))
            {
                switch (value)
                {
                    case "default":
                        return (T)(object)XML_SPACE.DEFAULT;
                    case "preserve":
                        return (T)(object)XML_SPACE.PRESERVE;
                    case "":
                        return (T)(object)XML_SPACE.NONE;
                }
            }
            else if (currType == typeof(PAGE_ORIENTATION))
            {
                switch (value)
                {
                    case "portrait":
                        return (T)(object)PAGE_ORIENTATION.PORTRAIT;
                    case "landscape":
                        return (T)(object)PAGE_ORIENTATION.LANSCAPE;
                    case "":
                        return (T)(object)PAGE_ORIENTATION.NONE;
                }
            }
            else if (currType == typeof(NUM_FMT))
            {
                switch (value)
                {
                    case "":
                        return (T)(object)NUM_FMT.DEFAULT;
                    case "upperLetter":
                        return (T)(object)NUM_FMT.UPPER_LETTER;
                    case "lowerRoman":
                        return (T)(object)NUM_FMT.LOWER_ROMAN;
                    case "upperRoman":
                        return (T)(object)NUM_FMT.UPPER_ROMAN;
                    case "chicago":
                        return (T)(object)NUM_FMT.CHICAGO;
                    case "bullet":
                        return (T)(object)NUM_FMT.BULLET;
                    case "cardinalText":
                        return (T)(object)NUM_FMT.CARDINAL_TEXT;
                    case "decimal":
                        return (T)(object)NUM_FMT.DECIMAL;
                    case "decimalEnclosedCircle":
                        return (T)(object)NUM_FMT.DECIMAL_ENCLOSED_CIRCLE;
                    case "decimalEnclosedFullstop":
                        return (T)(object)NUM_FMT.DECIMAL_ENCLOSED_FULL_STOP;
                    case "decimalEnclosedParen":
                        return (T)(object)NUM_FMT.DECIMAL_ENCLOSED_PAREN;
                    case "decimalZero":
                        return (T)(object)NUM_FMT.DECIMAL_ZERO;
                    case "none":
                        return (T)(object)NUM_FMT.NONE;
                    case "ordinalText":
                        return (T)(object)NUM_FMT.ORDINAL_TEXT;
                    case "ideographDigital":
                        return (T)(object)NUM_FMT.IDEOGRAPN_DIGITAL;

                }
            }
            else if (currType == typeof(FLD_CHAR_TYPE))
            {
                switch (value)
                {
                    case "begin":
                        return (T)(object)FLD_CHAR_TYPE.BEGIN;
                    case "separate":
                        return (T)(object)FLD_CHAR_TYPE.SEPARATE;
                    case "end":
                        return (T)(object)FLD_CHAR_TYPE.END;
                }
            }
            else if (currType == typeof(FOOTER_TYPE))
            {
                switch (value)
                {
                    case "":
                        return (T)(object)FOOTER_TYPE.NONE;
                    case "separator":
                        return (T)(object)FOOTER_TYPE.SEPARATOR;
                    case "continuationSeparator":
                        return (T)(object)FOOTER_TYPE.CONTINUATION_SEPAPRATOR;
                }
            }
            else if (currType == typeof(DOC_PART_GALLERY_VALUE))
            {
                switch (value)
                {
                    case "":
                        return (T)(object)DOC_PART_GALLERY_VALUE.NONE;
                    case "Page Numbers (Bottom of Page)":
                        return (T)(object)DOC_PART_GALLERY_VALUE.PAGE_NUMBERS_BOTTOM_OF_PAGE;
                    case "Page Numbers (Top of Page)":
                        return (T)(object)DOC_PART_GALLERY_VALUE.PAGE_NUMBERS_TOP_OF_PAGE;
                }
            }
            else if (currType == typeof(REFERENCE_TYPE))
            {
                switch (value)
                {
                    case "first":
                        return (T)(object)REFERENCE_TYPE.EVEN;
                    case "even":
                        return (T)(object)REFERENCE_TYPE.EVEN;
                    case "default":
                        return (T)(object)REFERENCE_TYPE.DEFAULT;
                }
            }
            else if (currType == typeof(HORIZONTAL_ALIGN))
            {
                switch (value)
                {
                    case "left":
                        return (T)(object)HORIZONTAL_ALIGN.LEFT;
                    case "center":
                        return (T)(object)HORIZONTAL_ALIGN.CENTER;
                    case "right":
                        return (T)(object)HORIZONTAL_ALIGN.RIGHT;
                    case "both":
                        return (T)(object)HORIZONTAL_ALIGN.BOTH;
                }
            }
            else if (currType == typeof(MULTI_LEVEL_TYPE))
            {
                switch (value.ToLower())
                {
                    case "singlelevel":
                        return (T)(object)MULTI_LEVEL_TYPE.SINGLE_LEVEL;
                    case "multilevel":
                        return (T)(object)MULTI_LEVEL_TYPE.MULTI_LEVEL;
                    case "hybridmultilevel":
                        return (T)(object)MULTI_LEVEL_TYPE.HYBRID_MULTY_LEVEL;
                }
            }
            else if (currType == typeof(BORDER))
            {
                switch (value)
                {
                    case "left":
                        return (T)(object)BORDER.LEFT;
                    case "right":
                        return (T)(object)BORDER.RIGHT;
                    case "top":
                        return (T)(object)BORDER.TOP;
                    case "bottom":
                        return (T)(object)BORDER.BOTTOM;
                    case "between":
                        return (T)(object)BORDER.BETWEEN;
                    case "bar":
                        return (T)(object)BORDER.BAR;
                    case "insideH":
                        return (T)(object)BORDER.INSIDE_H;
                    case "insideV":
                        return (T)(object)BORDER.INSIDE_V;
                }
            }
            throw new NotImplementedException();
        }
    }
}
