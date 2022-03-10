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
        /// <summary>
        /// одиночная линия
        /// </summary>
        SINGLE,
        /// <summary>
        /// линия с чередованием тонких и толстых штрихов
        /// </summary>
        DASH_DOT_STROKED,
        /// <summary>
        /// пунктирная линия
        /// </summary>
        DASHED,
        /// <summary>
        /// пунктирная линия с небольшими промежутками
        /// </summary>
        DASH_SMALL_GAP,
        /// <summary>
        /// линия с чередующимися точками и тире
        /// </summary>
        DOT_DASH,
        /// <summary>
        ///  линия с повторяющейся точкой
        /// </summary>
        DOT_DOT_DASH,
        /// <summary>
        /// пунктирная линия
        /// </summary>
        DOTTED,
        /// <summary>
        /// двойная линия
        /// </summary>
        DOUBLE,
        /// <summary>
        /// двойная волнистая линия
        /// </summary>
        DOUBLE_WAVE,
        /// <summary>
        /// набор линий
        /// </summary>
        INSET,
        /// <summary>
        /// без границ
        /// </summary>
        NIL,
        /// <summary>
        /// без границ
        /// </summary>
        NONE,
        /// <summary>
        /// начальный набор линий
        /// </summary>
        OUTSET,
        /// <summary>
        /// одна строка
        /// </summary>
        THICK,
        /// <summary>
        /// толстая линия, заключенная в тонкую линию с большим
        /// </summary>
        THICK_THIN_LARGE_GAP,
        /// <summary>
        /// толстая линия внутри тонкой линии со средним
        /// </summary>
        THICK_THIN_MEDIUM_GAP,
        /// <summary>
        /// толстая линия, содержащаяся внутри тонкой линии с небольшим промежуточным промежутком
        /// </summary>
        THICK_THIN_SMALL_GAP,
        /// <summary>
        /// тонкая линия, заключенная в толстую линию с большим
        /// </summary>
        THIN_THICK_LARGE_GAP,
        /// <summary>
        /// толстая линия внутри тонкой линии со средним
        /// </summary>
        THIN_THICK_MEDIUM_GAP,
        /// <summary>
        /// толстая линия, содержащаяся внутри тонкой линии с небольшим промежуточным промежутком
        /// </summary>
        THIN_THICK_SMALL_GAP,
        /// <summary>
        /// тонкий
        /// </summary>
        THIN_THICK_THIN_LARGE_GAP,
        /// <summary>
        /// тонкий
        /// </summary>
        THIN_THICK_THIN_MEDIUM_GAP,
        /// <summary>
        /// тонкий
        /// </summary>
        THIN_THICK_THIN_SMALL_GAP,
        /// <summary>
        /// три
        /// </summary>
        THREE_DEMBOSS,
        /// <summary>
        /// три
        /// </summary>
        THREE_DENGRAVE,
        /// <summary>
        /// тройная линия
        /// </summary>
        TRIPLE,
        /// <summary>
        /// волнистая линия 
        /// </summary>
        WAVE

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

    public enum RELATIONSIP_TYPE
    {
        FOOTER, STYLES, ENDNOTES, NUMBERING, CUSTOM_XML, FOOTNOTES, WEB_SETTINGS, THEME, SETTINGS, FONT_TABLE, HEADER, HYPERLINK,COMMENT
    }

    public enum ASCII_THEME
    {
        NONE,
        /// <summary>
        /// Указывает, что текущий шрифт является ссылкой на основной шрифт темы для диапазона ASCII.
        /// <summary>
        MAJOR_ASCII,
        /// <summary>
        /// Указывает, что текущий шрифт является ссылкой на основной шрифт темы для диапазона сложных сценариев.
        /// <summary>
        MAJOR_BIDI,
        /// <summary>
        /// Указывает, что текущий шрифт является ссылкой на основной шрифт темы для восточноазиатского диапазона.
        /// <summary>
        MAJOR_EAST_ASIA,
        /// <summary>
        /// Указывает, что текущий шрифт является ссылкой на основной шрифт темы для диапазона High ANSI.
        /// <summary>
        MAJOR_HANSI,
        /// <summary>
        /// Указывает, что текущий шрифт является ссылкой на дополнительный шрифт темы для диапазона ASCII.
        /// <summary>
        MINOR_ASCII,
        /// <summary>
        /// Указывает, что текущий шрифт является ссылкой на второстепенный шрифт темы для диапазона сложных сценариев.
        /// <summary>
        MINOR_BIDI,
        /// <summary>
        /// Указывает, что текущий шрифт является ссылкой на дополнительный шрифт темы для восточноазиатского диапазона.
        /// <summary>
        MINOR_EAST_ASIA,
        /// <summary>
        /// Указывает, что текущий шрифт является ссылкой на дополнительный шрифт темы для диапазона High ANSI.
        /// <summary>
        MINOR_HANSI
    }

    public enum NUM_FMT
    {
        /// <summary>
        /// AIUEO Орден Хирагана
        /// <summary>
        AIUEO,
        /// <summary>
        /// Полноразмерный AIUEO Заказать Хирагана
        /// <summary>
        AIUEO_FULL_WIDTH,
        /// <summary>
        /// Арабские цифры абджад
        /// <summary>
        ARABIC_ABJAD,
        /// <summary>
        /// Арабский алфавит
        /// <summary>
        ARABIC_ALPHA,
        /// <summary>
        /// Пуля
        /// <summary>
        BULLET,
        /// <summary>
        /// Кардинальный текст
        /// <summary>
        CARDINAL_TEXT,
        /// <summary>
        /// Чикагское руководство стиля
        /// <summary>
        CHICAGO,
        /// <summary>
        /// Китайская система подсчета
        /// <summary>
        CHINESE_COUNTING,
        /// <summary>
        /// Китайская система подсчета тысяч
        /// <summary>
        CHINESE_COUNTING_THOUSAND,
        /// <summary>
        /// Китайский упрощенный юридический формат
        /// <summary>
        CHINESE_LEGAL_SIMPLIFIED,
        /// <summary>
        /// Корейская нумерация Chosung
        /// <summary>
        CHOSUNG,
        /// <summary>
        /// Десятичные числа
        /// <summary>
        DECIMAL,
        /// <summary>
        /// Десятичные числа, заключенные в круг
        /// <summary>
        DECIMAL_ENCLOSED_CIRCLE,
        /// <summary>
        /// Десятичные числа, заключенные в круг
        /// <summary>
        DECIMAL_ENCLOSED_CIRCLE_CHINESE,
        /// <summary>
        /// Десятичные числа с точкой
        /// <summary>
        DECIMAL_ENCLOSED_FULLSTOP,
        /// <summary>
        /// Десятичные числа в круглых скобках
        /// <summary>
        DECIMAL_ENCLOSED_PAREN,
        /// <summary>
        /// Двухбайтовые арабские цифры
        /// <summary>
        DECIMAL_FULL_WIDTH,
        /// <summary>
        /// Альтернативные двухбайтовые арабские цифры
        /// <summary>
        DECIMAL_FULL_WIDTH2,
        /// <summary>
        /// Однобайтовые арабские цифры
        /// <summary>
        DECIMAL_HALF_WIDTH,
        /// <summary>
        /// Начальные нулевые арабские цифры
        /// <summary>
        DECIMAL_ZERO,
        /// <summary>
        /// Корейская нумерация Ганада
        /// <summary>
        GANADA,
        /// <summary>
        /// Еврейские цифры
        /// <summary>
        HEBREW1,
        /// <summary>
        /// Еврейский алфавит
        /// <summary>
        HEBREW2,
        /// <summary>
        /// Шестнадцатеричная нумерация
        /// <summary>
        HEX,
        /// <summary>
        /// Согласные хинди
        /// <summary>
        HINDI_CONSONANTS,
        /// <summary>
        /// Система подсчета хинди
        /// <summary>
        HINDI_COUNTING,
        /// <summary>
        /// Числа на хинди
        /// <summary>
        HINDI_NUMBERS,
        /// <summary>
        /// Гласные хинди
        /// <summary>
        HINDI_VOWELS,
        /// <summary>
        /// Идеографы
        /// <summary>
        IDEOGRAPH_DIGITAL,
        /// <summary>
        /// Идеографы, заключенные в круг
        /// <summary>
        IDEOGRAPH_ENCLOSED_CIRCLE,
        /// <summary>
        /// Формат традиционного правового идеографа
        /// <summary>
        IDEOGRAPH_LEGAL_TRADITIONAL,
        /// <summary>
        /// Традиционный формат идеографа
        /// <summary>
        IDEOGRAPH_TRADITIONAL,
        /// <summary>
        /// Формат идеографа зодиака
        /// <summary>
        IDEOGRAPH_ZODIAC,
        /// <summary>
        /// Традиционный формат идеографа зодиака
        /// <summary>
        IDEOGRAPH_ZODIAC_TRADITIONAL,
        /// <summary>
        /// Ироха Заказанная Катакана
        /// <summary>
        IROHA,
        /// <summary>
        /// Полноразмерная Ироха Заказанная Катакана
        /// <summary>
        IROHA_FULL_WIDTH,
        /// <summary>
        /// Японская система подсчета
        /// <summary>
        JAPANESE_COUNTING,
        /// <summary>
        /// Японская цифровая система подсчета десяти тысяч
        /// <summary>
        JAPANESE_DIGITAL_TEN_THOUSAND,
        /// <summary>
        /// Японская юридическая нумерация
        /// <summary>
        JAPANESE_LEGAL,
        /// <summary>
        /// Корейская система подсчета
        /// <summary>
        KOREAN_COUNTING,
        /// <summary>
        /// Корейская цифровая система подсчета
        /// <summary>
        KOREAN_DIGITAL,
        /// <summary>
        /// Альтернативная корейская цифровая система подсчета
        /// <summary>
        KOREAN_DIGITAL2,
        /// <summary>
        /// Корейская юридическая нумерация
        /// <summary>
        KOREAN_LEGAL,
        /// <summary>
        /// Строчный латинский алфавит
        /// <summary>
        LOWER_LETTER,
        /// <summary>
        /// Строчные римские цифры
        /// <summary>
        LOWER_ROMAN,
        /// <summary>
        /// Без нумерации
        /// <summary>
        NONE,
        /// <summary>
        /// Число с тире
        /// <summary>
        NUMBERIN_DASH,
        /// <summary>
        /// Порядковый
        /// <summary>
        ORDINAL,
        /// <summary>
        /// Порядковый текст
        /// <summary>
        ORDINAL_TEXT,
        /// <summary>
        /// Русский алфавит в нижнем регистре
        /// <summary>
        RUSSIAN_LOWER,
        /// <summary>
        /// Русский алфавит в верхнем регистре
        /// <summary>
        RUSSIAN_UPPER,
        /// <summary>
        /// Тайваньская система подсчета
        /// <summary>
        TAIWANESE_COUNTING,
        /// <summary>
        /// Тайваньская система подсчета тысяч
        /// <summary>
        TAIWANESE_COUNTING_THOUSAND,
        /// <summary>
        /// Тайваньская цифровая система подсчета
        /// <summary>
        TAIWANESE_DIGITAL,
        /// <summary>
        /// Тайская система подсчета
        /// <summary>
        THAI_COUNTING,
        /// <summary>
        /// Тайские буквы
        /// <summary>
        THAI_LETTERS,
        /// <summary>
        /// Тайские цифры
        /// <summary>
        THAI_NUMBERS,
        /// <summary>
        /// Латинский алфавит в верхнем регистре
        /// <summary>
        UPPER_LETTER,
        /// <summary>
        /// Римские цифры в верхнем регистре
        /// <summary>
        UPPER_ROMAN,
        /// <summary>
        /// Вьетнамские цифры
        /// <summary>
        VIETNAMESE_COUNTING
    }

    public enum BR_TYPE
    {
        NONE,
        COLUMN,
        PAGE,
        TEXT_WRAPPING
    }

    public enum VERTICAL_ALIGN { NONE,TOP, CENTER, BOTTOM, BOTH }

    public enum SECT_TYPE { NONE,
        /// <summary>
        /// Begins the section on the next paragraph. Certain page-level section properties cannot be specified, as they are inherited from the previous section. If a footnote occurs of the same page as a section of this kind, the new section begins on the following page.
        /// </summary>
        CONTINUOUS,
        /// <summary>
        /// The section begins on the next even-numbered page, leaving the next odd page blank if necessary.
        /// </summary>
        EVEN_PAGE,
        /// <summary>
        ///  The section begins on the following column on the page.
        /// </summary>
        NEXT_COLUMN,
        /// <summary>
        /// The section begins on the following page.
        /// </summary>
        NEXT_PAGE,
        /// <summary>
        /// The section begins on the next odd-numbered page, leaving the next even page blank if necessary.
        /// </summary>
        ODD_PAGE
    }

    public static class EnumExtentions
    {
        public static string ToStringValue(this LINE_TYPE e) {
            switch (e)
            {
                case LINE_TYPE.SINGLE:
                    return "single";
                case LINE_TYPE.DASH_DOT_STROKED:
                    return "dashDotStroked";
                case LINE_TYPE.DASHED:
                    return "dashed";
                case LINE_TYPE.DASH_SMALL_GAP:
                    return "dashSmallGap";
                case LINE_TYPE.DOT_DASH:
                    return "dotDash";
                case LINE_TYPE.DOT_DOT_DASH:
                    return "dotDotDash";
                case LINE_TYPE.DOTTED:
                    return "dotted";
                case LINE_TYPE.DOUBLE:
                    return "double";
                case LINE_TYPE.DOUBLE_WAVE:
                    return "doubleWave";
                case LINE_TYPE.INSET:
                    return "inset";
                case LINE_TYPE.NIL:
                    return "nil";
                case LINE_TYPE.NONE:
                    return "none";
                case LINE_TYPE.OUTSET:
                    return "outset";
                case LINE_TYPE.THICK:
                    return "thick";
                case LINE_TYPE.THICK_THIN_LARGE_GAP:
                    return "thickThinLargeGap";
                case LINE_TYPE.THICK_THIN_MEDIUM_GAP:
                    return "thickThinMediumGap";
                case LINE_TYPE.THICK_THIN_SMALL_GAP:
                    return "thickThinSmallGap";
                case LINE_TYPE.THIN_THICK_LARGE_GAP:
                    return "thinThickLargeGap";
                case LINE_TYPE.THIN_THICK_MEDIUM_GAP:
                    return "thinThickMediumGap";
                case LINE_TYPE.THIN_THICK_SMALL_GAP:
                    return "thinThickSmallGap";
                case LINE_TYPE.THIN_THICK_THIN_LARGE_GAP:
                    return "thinThickThinLargeGap";
                case LINE_TYPE.THIN_THICK_THIN_MEDIUM_GAP:
                    return "thinThickThinMediumGap";
                case LINE_TYPE.THIN_THICK_THIN_SMALL_GAP:
                    return "thinThickThinSmallGap";
                case LINE_TYPE.THREE_DEMBOSS:
                    return "threeDEmboss";
                case LINE_TYPE.THREE_DENGRAVE:
                    return "threeDEngrave";
                case LINE_TYPE.TRIPLE:
                    return "triple";
                case LINE_TYPE.WAVE:
                    return "wave";
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
        //https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_NumberFormat_topic_ID0EDNB3.html#topic_ID0EDNB3
        public static string ToStringValue(this NUM_FMT e)
        {
            switch (e)
            {
                case NUM_FMT.AIUEO:
                    return "aiueo";
                case NUM_FMT.AIUEO_FULL_WIDTH:
                    return "aiueoFullWidth";
                case NUM_FMT.ARABIC_ABJAD:
                    return "arabicAbjad";
                case NUM_FMT.ARABIC_ALPHA:
                    return "arabicAlpha";
                case NUM_FMT.BULLET:
                    return "bullet";
                case NUM_FMT.CARDINAL_TEXT:
                    return "cardinalText";
                case NUM_FMT.CHICAGO:
                    return "chicago";
                case NUM_FMT.CHINESE_COUNTING:
                    return "chineseCounting";
                case NUM_FMT.CHINESE_COUNTING_THOUSAND:
                    return "chineseCountingThousand";
                case NUM_FMT.CHINESE_LEGAL_SIMPLIFIED:
                    return "chineseLegalSimplified";
                case NUM_FMT.CHOSUNG:
                    return "chosung";
                case NUM_FMT.DECIMAL:
                    return "decimal";
                case NUM_FMT.DECIMAL_ENCLOSED_CIRCLE:
                    return "decimalEnclosedCircle";
                case NUM_FMT.DECIMAL_ENCLOSED_CIRCLE_CHINESE:
                    return "decimalEnclosedCircleChinese";
                case NUM_FMT.DECIMAL_ENCLOSED_FULLSTOP:
                    return "decimalEnclosedFullstop";
                case NUM_FMT.DECIMAL_ENCLOSED_PAREN:
                    return "decimalEnclosedParen";
                case NUM_FMT.DECIMAL_FULL_WIDTH:
                    return "decimalFullWidth";
                case NUM_FMT.DECIMAL_FULL_WIDTH2:
                    return "decimalFullWidth2";
                case NUM_FMT.DECIMAL_HALF_WIDTH:
                    return "decimalHalfWidth";
                case NUM_FMT.DECIMAL_ZERO:
                    return "decimalZero";
                case NUM_FMT.GANADA:
                    return "ganada";
                case NUM_FMT.HEBREW1:
                    return "hebrew1";
                case NUM_FMT.HEBREW2:
                    return "hebrew2";
                case NUM_FMT.HEX:
                    return "hex";
                case NUM_FMT.HINDI_CONSONANTS:
                    return "hindiConsonants";
                case NUM_FMT.HINDI_COUNTING:
                    return "hindiCounting";
                case NUM_FMT.HINDI_NUMBERS:
                    return "hindiNumbers";
                case NUM_FMT.HINDI_VOWELS:
                    return "hindiVowels";
                case NUM_FMT.IDEOGRAPH_DIGITAL:
                    return "ideographDigital";
                case NUM_FMT.IDEOGRAPH_ENCLOSED_CIRCLE:
                    return "ideographEnclosedCircle";
                case NUM_FMT.IDEOGRAPH_LEGAL_TRADITIONAL:
                    return "ideographLegalTraditional";
                case NUM_FMT.IDEOGRAPH_TRADITIONAL:
                    return "ideographTraditional";
                case NUM_FMT.IDEOGRAPH_ZODIAC:
                    return "ideographZodiac";
                case NUM_FMT.IDEOGRAPH_ZODIAC_TRADITIONAL:
                    return "ideographZodiacTraditional";
                case NUM_FMT.IROHA:
                    return "iroha";
                case NUM_FMT.IROHA_FULL_WIDTH:
                    return "irohaFullWidth";
                case NUM_FMT.JAPANESE_COUNTING:
                    return "japaneseCounting";
                case NUM_FMT.JAPANESE_DIGITAL_TEN_THOUSAND:
                    return "japaneseDigitalTenThousand";
                case NUM_FMT.JAPANESE_LEGAL:
                    return "japaneseLegal";
                case NUM_FMT.KOREAN_COUNTING:
                    return "koreanCounting";
                case NUM_FMT.KOREAN_DIGITAL:
                    return "koreanDigital";
                case NUM_FMT.KOREAN_DIGITAL2:
                    return "koreanDigital2";
                case NUM_FMT.KOREAN_LEGAL:
                    return "koreanLegal";
                case NUM_FMT.LOWER_LETTER:
                    return "lowerLetter";
                case NUM_FMT.LOWER_ROMAN:
                    return "lowerRoman";
                case NUM_FMT.NONE:
                    return "none";
                case NUM_FMT.NUMBERIN_DASH:
                    return "numberInDash";
                case NUM_FMT.ORDINAL:
                    return "ordinal";
                case NUM_FMT.ORDINAL_TEXT:
                    return "ordinalText";
                case NUM_FMT.RUSSIAN_LOWER:
                    return "russianLower";
                case NUM_FMT.RUSSIAN_UPPER:
                    return "russianUpper";
                case NUM_FMT.TAIWANESE_COUNTING:
                    return "taiwaneseCounting";
                case NUM_FMT.TAIWANESE_COUNTING_THOUSAND:
                    return "taiwaneseCountingThousand";
                case NUM_FMT.TAIWANESE_DIGITAL:
                    return "taiwaneseDigital";
                case NUM_FMT.THAI_COUNTING:
                    return "thaiCounting";
                case NUM_FMT.THAI_LETTERS:
                    return "thaiLetters";
                case NUM_FMT.THAI_NUMBERS:
                    return "thaiNumbers";
                case NUM_FMT.UPPER_LETTER:
                    return "upperLetter";
                case NUM_FMT.UPPER_ROMAN:
                    return "upperRoman";
                case NUM_FMT.VIETNAMESE_COUNTING:
                    return "vietnameseCounting";
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
        public static string ToStringValue(this BORDER_TYPE e)
        {
            switch (e)
            {
                case BORDER_TYPE.LEFT:
                    return "left";
                case BORDER_TYPE.RIGHT:
                    return "right";
                case BORDER_TYPE.TOP:
                    return "top";
                case BORDER_TYPE.BOTTOM:
                    return "bottom";
                case BORDER_TYPE.BETWEEN:
                    return "between";
                case BORDER_TYPE.BAR:
                    return "bar";
                case BORDER_TYPE.INSIDE_H:
                    return "insideH";
                case BORDER_TYPE.INSIDE_V:
                    return "insideV";
            }
            throw new NotImplementedException();
        }
        public static string ToStringValue(this ASCII_THEME e) { switch (e) {
                case ASCII_THEME.NONE:
                    return "";
                case ASCII_THEME.MAJOR_ASCII:
                    return "majorAscii";
                case ASCII_THEME.MAJOR_BIDI:
                    return "majorBidi";
                case ASCII_THEME.MAJOR_EAST_ASIA:
                    return "majorEastAsia";
                case ASCII_THEME.MAJOR_HANSI:
                    return "majorHAnsi";
                case ASCII_THEME.MINOR_ASCII:
                    return "minorAscii";
                case ASCII_THEME.MINOR_BIDI:
                    return "minorBidi";
                case ASCII_THEME.MINOR_EAST_ASIA:
                    return "minorEastAsia";
                case ASCII_THEME.MINOR_HANSI:
                    return "minorHAnsi";
            }
            throw new NotImplementedException();
        }
        public static string ToStringValue(this RELATIONSIP_TYPE e)
        {
            switch (e)
            {
                case RELATIONSIP_TYPE.FOOTER:
                    return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
                case RELATIONSIP_TYPE.STYLES:
                    return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
                case RELATIONSIP_TYPE.ENDNOTES:
                    return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes";
                case RELATIONSIP_TYPE.NUMBERING:
                    return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
                case RELATIONSIP_TYPE.CUSTOM_XML:
                    return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
                case RELATIONSIP_TYPE.FOOTNOTES:
                    return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes";
                case RELATIONSIP_TYPE.WEB_SETTINGS:
                    return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings";
                case RELATIONSIP_TYPE.THEME:
                    return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
                case RELATIONSIP_TYPE.SETTINGS:
                    return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
                case RELATIONSIP_TYPE.FONT_TABLE:
                    return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable";
                case RELATIONSIP_TYPE.HEADER:
                    return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
                case RELATIONSIP_TYPE.HYPERLINK:
                    return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
                case RELATIONSIP_TYPE.COMMENT:
                    return "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
            }
            throw new NotImplementedException();
        }

        public static string ToStringValue(this BR_TYPE e)
        {
            switch (e)
            {
                case BR_TYPE.NONE:
                    return "";
                case BR_TYPE.COLUMN:
                    return "column";
                case BR_TYPE.PAGE:
                    return "page";
                case BR_TYPE.TEXT_WRAPPING:
                    return "textWrapping";
                default:
                    throw new ArgumentOutOfRangeException(nameof(e), e, null);
            }
            throw new NotImplementedException();
        }
        public static string ToStringValue(this VERTICAL_ALIGN e)
        {
            switch (e)
            {
                case VERTICAL_ALIGN.BOTH:
                    return "Both";
                case VERTICAL_ALIGN.BOTTOM:
                    return "Bottom";
                case VERTICAL_ALIGN.CENTER:
                    return "Center";
                case VERTICAL_ALIGN.TOP:
                    return "Top";
                default:
                    throw new ArgumentOutOfRangeException(nameof(e), e, null);
            }
            throw new NotImplementedException();
        }
        public static string ToStringValue(this SECT_TYPE e)
        {
            switch (e)
            {
                case SECT_TYPE.NONE:
                    return "";
                case SECT_TYPE.CONTINUOUS:
                    return "continuous";
                case SECT_TYPE.EVEN_PAGE:
                    return "evenPage";
                case SECT_TYPE.NEXT_COLUMN:
                    return "nextColumn";
                case SECT_TYPE.NEXT_PAGE:
                    return "nextPage";
                case SECT_TYPE.ODD_PAGE:
                    return "oddPage";
                default:
                    throw new ArgumentOutOfRangeException(nameof(e), e, null);
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
                    case "single": return (T)(object)LINE_TYPE.SINGLE;
                    case "dashDotStroked": return (T)(object)LINE_TYPE.DASH_DOT_STROKED;
                    case "dashed": return (T)(object)LINE_TYPE.DASHED;
                    case "dashSmallGap": return (T)(object)LINE_TYPE.DASH_SMALL_GAP;
                    case "dotDash": return (T)(object)LINE_TYPE.DOT_DASH;
                    case "dotDotDash": return (T)(object)LINE_TYPE.DOT_DOT_DASH;
                    case "dotted": return (T)(object)LINE_TYPE.DOTTED;
                    case "double": return (T)(object)LINE_TYPE.DOUBLE;
                    case "doubleWave": return (T)(object)LINE_TYPE.DOUBLE_WAVE;
                    case "inset": return (T)(object)LINE_TYPE.INSET;
                    case "nil": return (T)(object)LINE_TYPE.NIL;
                    case "none": return (T)(object)LINE_TYPE.NONE;
                    case "outset": return (T)(object)LINE_TYPE.OUTSET;
                    case "thick": return (T)(object)LINE_TYPE.THICK;
                    case "thickThinLargeGap": return (T)(object)LINE_TYPE.THICK_THIN_LARGE_GAP;
                    case "thickThinMediumGap": return (T)(object)LINE_TYPE.THICK_THIN_MEDIUM_GAP;
                    case "thickThinSmallGap": return (T)(object)LINE_TYPE.THICK_THIN_SMALL_GAP;
                    case "thinThickLargeGap": return (T)(object)LINE_TYPE.THIN_THICK_LARGE_GAP;
                    case "thinThickMediumGap": return (T)(object)LINE_TYPE.THIN_THICK_MEDIUM_GAP;
                    case "thinThickSmallGap": return (T)(object)LINE_TYPE.THIN_THICK_SMALL_GAP;
                    case "thinThickThinLargeGap": return (T)(object)LINE_TYPE.THIN_THICK_THIN_LARGE_GAP;
                    case "thinThickThinMediumGap": return (T)(object)LINE_TYPE.THIN_THICK_THIN_MEDIUM_GAP;
                    case "thinThickThinSmallGap": return (T)(object)LINE_TYPE.THIN_THICK_THIN_SMALL_GAP;
                    case "threeDEmboss": return (T)(object)LINE_TYPE.THREE_DEMBOSS;
                    case "threeDEngrave": return (T)(object)LINE_TYPE.THREE_DENGRAVE;
                    case "triple": return (T)(object)LINE_TYPE.TRIPLE;
                    case "wave": return (T)(object)LINE_TYPE.WAVE;
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
                    case "aiueo": return (T)(object)NUM_FMT.AIUEO;
                    case "aiueoFullWidth": return (T)(object)NUM_FMT.AIUEO_FULL_WIDTH;
                    case "arabicAbjad": return (T)(object)NUM_FMT.ARABIC_ABJAD;
                    case "arabicAlpha": return (T)(object)NUM_FMT.ARABIC_ALPHA;
                    case "bullet": return (T)(object)NUM_FMT.BULLET;
                    case "cardinalText": return (T)(object)NUM_FMT.CARDINAL_TEXT;
                    case "chicago": return (T)(object)NUM_FMT.CHICAGO;
                    case "chineseCounting": return (T)(object)NUM_FMT.CHINESE_COUNTING;
                    case "chineseCountingThousand": return (T)(object)NUM_FMT.CHINESE_COUNTING_THOUSAND;
                    case "chineseLegalSimplified": return (T)(object)NUM_FMT.CHINESE_LEGAL_SIMPLIFIED;
                    case "chosung": return (T)(object)NUM_FMT.CHOSUNG;
                    case "decimal": return (T)(object)NUM_FMT.DECIMAL;
                    case "decimalEnclosedCircle": return (T)(object)NUM_FMT.DECIMAL_ENCLOSED_CIRCLE;
                    case "decimalEnclosedCircleChinese": return (T)(object)NUM_FMT.DECIMAL_ENCLOSED_CIRCLE_CHINESE;
                    case "decimalEnclosedFullstop": return (T)(object)NUM_FMT.DECIMAL_ENCLOSED_FULLSTOP;
                    case "decimalEnclosedParen": return (T)(object)NUM_FMT.DECIMAL_ENCLOSED_PAREN;
                    case "decimalFullWidth": return (T)(object)NUM_FMT.DECIMAL_FULL_WIDTH;
                    case "decimalFullWidth2": return (T)(object)NUM_FMT.DECIMAL_FULL_WIDTH2;
                    case "decimalHalfWidth": return (T)(object)NUM_FMT.DECIMAL_HALF_WIDTH;
                    case "decimalZero": return (T)(object)NUM_FMT.DECIMAL_ZERO;
                    case "ganada": return (T)(object)NUM_FMT.GANADA;
                    case "hebrew1": return (T)(object)NUM_FMT.HEBREW1;
                    case "hebrew2": return (T)(object)NUM_FMT.HEBREW2;
                    case "hex": return (T)(object)NUM_FMT.HEX;
                    case "hindiConsonants": return (T)(object)NUM_FMT.HINDI_CONSONANTS;
                    case "hindiCounting": return (T)(object)NUM_FMT.HINDI_COUNTING;
                    case "hindiNumbers": return (T)(object)NUM_FMT.HINDI_NUMBERS;
                    case "hindiVowels": return (T)(object)NUM_FMT.HINDI_VOWELS;
                    case "ideographDigital": return (T)(object)NUM_FMT.IDEOGRAPH_DIGITAL;
                    case "ideographEnclosedCircle": return (T)(object)NUM_FMT.IDEOGRAPH_ENCLOSED_CIRCLE;
                    case "ideographLegalTraditional": return (T)(object)NUM_FMT.IDEOGRAPH_LEGAL_TRADITIONAL;
                    case "ideographTraditional": return (T)(object)NUM_FMT.IDEOGRAPH_TRADITIONAL;
                    case "ideographZodiac": return (T)(object)NUM_FMT.IDEOGRAPH_ZODIAC;
                    case "ideographZodiacTraditional": return (T)(object)NUM_FMT.IDEOGRAPH_ZODIAC_TRADITIONAL;
                    case "iroha": return (T)(object)NUM_FMT.IROHA;
                    case "irohaFullWidth": return (T)(object)NUM_FMT.IROHA_FULL_WIDTH;
                    case "japaneseCounting": return (T)(object)NUM_FMT.JAPANESE_COUNTING;
                    case "japaneseDigitalTenThousand": return (T)(object)NUM_FMT.JAPANESE_DIGITAL_TEN_THOUSAND;
                    case "japaneseLegal": return (T)(object)NUM_FMT.JAPANESE_LEGAL;
                    case "koreanCounting": return (T)(object)NUM_FMT.KOREAN_COUNTING;
                    case "koreanDigital": return (T)(object)NUM_FMT.KOREAN_DIGITAL;
                    case "koreanDigital2": return (T)(object)NUM_FMT.KOREAN_DIGITAL2;
                    case "koreanLegal": return (T)(object)NUM_FMT.KOREAN_LEGAL;
                    case "lowerLetter": return (T)(object)NUM_FMT.LOWER_LETTER;
                    case "lowerRoman": return (T)(object)NUM_FMT.LOWER_ROMAN;
                    case "none": return (T)(object)NUM_FMT.NONE;
                    case "numberInDash": return (T)(object)NUM_FMT.NUMBERIN_DASH;
                    case "ordinal": return (T)(object)NUM_FMT.ORDINAL;
                    case "ordinalText": return (T)(object)NUM_FMT.ORDINAL_TEXT;
                    case "russianLower": return (T)(object)NUM_FMT.RUSSIAN_LOWER;
                    case "russianUpper": return (T)(object)NUM_FMT.RUSSIAN_UPPER;
                    case "taiwaneseCounting": return (T)(object)NUM_FMT.TAIWANESE_COUNTING;
                    case "taiwaneseCountingThousand": return (T)(object)NUM_FMT.TAIWANESE_COUNTING_THOUSAND;
                    case "taiwaneseDigital": return (T)(object)NUM_FMT.TAIWANESE_DIGITAL;
                    case "thaiCounting": return (T)(object)NUM_FMT.THAI_COUNTING;
                    case "thaiLetters": return (T)(object)NUM_FMT.THAI_LETTERS;
                    case "thaiNumbers": return (T)(object)NUM_FMT.THAI_NUMBERS;
                    case "upperLetter": return (T)(object)NUM_FMT.UPPER_LETTER;
                    case "upperRoman": return (T)(object)NUM_FMT.UPPER_ROMAN;
                    case "vietnameseCounting": return (T)(object)NUM_FMT.VIETNAMESE_COUNTING;
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
                        return (T)(object)REFERENCE_TYPE.FIRST;
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
            else if (currType == typeof(BORDER_TYPE))
            {
                switch (value)
                {
                    case "left":
                        return (T)(object)BORDER_TYPE.LEFT;
                    case "right":
                        return (T)(object)BORDER_TYPE.RIGHT;
                    case "top":
                        return (T)(object)BORDER_TYPE.TOP;
                    case "bottom":
                        return (T)(object)BORDER_TYPE.BOTTOM;
                    case "between":
                        return (T)(object)BORDER_TYPE.BETWEEN;
                    case "bar":
                        return (T)(object)BORDER_TYPE.BAR;
                    case "insideH":
                        return (T)(object)BORDER_TYPE.INSIDE_H;
                    case "insideV":
                        return (T)(object)BORDER_TYPE.INSIDE_V;
                }
            }
            else if (currType == typeof(ASCII_THEME))
            {
                switch (value)
                {
                    case null:
                    case "":
                        return (T)(object)ASCII_THEME.NONE;
                    case "majorAscii": return (T)(object)ASCII_THEME.MAJOR_ASCII;
                    case "majorBidi": return (T)(object)ASCII_THEME.MAJOR_BIDI;
                    case "majorEastAsia": return (T)(object)ASCII_THEME.MAJOR_EAST_ASIA;
                    case "majorHAnsi": return (T)(object)ASCII_THEME.MAJOR_HANSI;
                    case "minorAscii": return (T)(object)ASCII_THEME.MINOR_ASCII;
                    case "minorBidi": return (T)(object)ASCII_THEME.MINOR_BIDI;
                    case "minorEastAsia": return (T)(object)ASCII_THEME.MINOR_EAST_ASIA;
                    case "minorHAnsi": return (T)(object)ASCII_THEME.MINOR_HANSI;
                }
            }
            else if (currType == typeof(RELATIONSIP_TYPE))
            {
                switch (value)
                {
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer":
                        return (T)(object)RELATIONSIP_TYPE.FOOTER;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles":
                        return (T)(object)RELATIONSIP_TYPE.STYLES;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes":
                        return (T)(object)RELATIONSIP_TYPE.ENDNOTES;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering":
                        return (T)(object)RELATIONSIP_TYPE.NUMBERING;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml":
                        return (T)(object)RELATIONSIP_TYPE.CUSTOM_XML;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes":
                        return (T)(object)RELATIONSIP_TYPE.FOOTNOTES;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings":
                        return (T)(object)RELATIONSIP_TYPE.WEB_SETTINGS;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme":
                        return (T)(object)RELATIONSIP_TYPE.THEME;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings":
                        return (T)(object)RELATIONSIP_TYPE.SETTINGS;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable":
                        return (T)(object)RELATIONSIP_TYPE.FONT_TABLE;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header":
                        return (T)(object)RELATIONSIP_TYPE.HEADER;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink":
                        return (T)(object)RELATIONSIP_TYPE.HYPERLINK;
                    case "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments":
                        return (T)(object)RELATIONSIP_TYPE.COMMENT;
                }
            }
            else if (currType == typeof(BR_TYPE))
            {
                switch (value)
                {
                    case "column":
                        return (T)(object)BR_TYPE.COLUMN;
                    case "page":
                        return (T)(object)BR_TYPE.PAGE;
                    case "textWrapping":
                        return (T)(object)BR_TYPE.TEXT_WRAPPING;
                }
            }
            else if (currType == typeof(VERTICAL_ALIGN))
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
            else if (currType == typeof(SECT_TYPE))
            {
                switch (value)
                {
                    case "continuous":
                        return (T)(object)SECT_TYPE.CONTINUOUS;
                    case "evenPage":
                        return (T)(object)SECT_TYPE.EVEN_PAGE;
                    case "nextColumn":
                        return (T)(object)SECT_TYPE.NEXT_COLUMN;
                    case "nextPage":
                        return (T)(object)SECT_TYPE.NEXT_PAGE;
                    case "oddPage":
                        return (T)(object)SECT_TYPE.ODD_PAGE;
                }
            }
            throw new NotImplementedException();
        }
    }


    public enum QUOTES
    {
        /// <summary>
        /// "двойные"
        /// </summary>
        DEFAULT,
        /// <summary>
        /// «ёлочки»
        /// </summary>
        FRANCE,
        /// <summary>
        /// ‟лапки”
        /// </summary>
        FRANCE_PAWS,
        /// <summary>
        /// ‹одиночные›
        /// </summary>
        FRANCE_SINGLE,
        /// <summary>
        /// „лапки“
        /// </summary>
        GERMAN,
        /// <summary>
        /// “английские двойные”
        /// </summary>
        ENGLISH_DOUBLE,
        /// <summary>
        /// ‘английские одиночные’
        /// </summary>
        ENGLISH_SINGLE
    }

    public static class QuoteExtentions
    { 
        public static Pair<char, char> ToQuotes (this QUOTES e)
        {
            switch (e)
            {
                case QUOTES.DEFAULT:
                    return new Pair<char, char>('"', '"');
                case QUOTES.FRANCE:
                    return new Pair<char, char>('«', '»');
                case QUOTES.FRANCE_PAWS:
                    return new Pair<char, char>('‟', '”');
                case QUOTES.FRANCE_SINGLE:
                    return new Pair<char, char>('‹', '›');
                case QUOTES.GERMAN:
                    return new Pair<char, char>('„', '“');
                case QUOTES.ENGLISH_DOUBLE:
                    return new Pair<char, char>('“', '”');
                case QUOTES.ENGLISH_SINGLE:
                    return new Pair<char, char>('‘', '’');

            }
            throw new NotImplementedException();
        }
    }
    
}
