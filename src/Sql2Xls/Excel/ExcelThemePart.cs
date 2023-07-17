using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Text;

namespace Sql2Xls.Excel
{
    public class ExcelThemePart : ExcelPart
    {
        protected const string themeRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";

        public ExcelThemePart(SpreadsheetDocument document, string relationshipId, ExcelExportContext context)
            : base(document, relationshipId, context)
        {
        }

        public ThemePart CreateThemePart(WorkbookPart workbookPart)
        {
            ThemePart themePart = workbookPart.AddNewPart<ThemePart>(RelationshipId);
            themePart.Theme = GenerateDefaultTheme();
            themePart.Theme.Save();

            if (Context.CanUseRelativePaths)
            {
                RelationshipId = ExcelHelper.UpdateWorkbookRelationshipsPath(Document, themePart, themeRelationshipType);
            }

            return themePart;
        }

        private Theme GenerateDefaultTheme()
        {
            var theme1 = new Theme { Name = "Office Theme" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            var themeElements1 = new ThemeElements();

            var colorScheme1 = new ColorScheme { Name = "Office" };

            var dark1Color1 = new Dark1Color();
            var systemColor1 = new SystemColor { Val = SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            var light1Color1 = new Light1Color();
            var systemColor2 = new SystemColor { Val = SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            var dark2Color1 = new Dark2Color();
            var rgbColorModelHex1 = new RgbColorModelHex { Val = "1F497D" };

            dark2Color1.Append(rgbColorModelHex1);

            var light2Color1 = new Light2Color();
            var rgbColorModelHex2 = new RgbColorModelHex { Val = "EEECE1" };

            light2Color1.Append(rgbColorModelHex2);

            var accent1Color1 = new Accent1Color();
            var rgbColorModelHex3 = new RgbColorModelHex { Val = "4F81BD" };

            accent1Color1.Append(rgbColorModelHex3);

            var accent2Color1 = new Accent2Color();
            var rgbColorModelHex4 = new RgbColorModelHex { Val = "C0504D" };

            accent2Color1.Append(rgbColorModelHex4);

            var accent3Color1 = new Accent3Color();
            var rgbColorModelHex5 = new RgbColorModelHex { Val = "9BBB59" };

            accent3Color1.Append(rgbColorModelHex5);

            var accent4Color1 = new Accent4Color();
            var rgbColorModelHex6 = new RgbColorModelHex { Val = "8064A2" };

            accent4Color1.Append(rgbColorModelHex6);

            var accent5Color1 = new Accent5Color();
            var rgbColorModelHex7 = new RgbColorModelHex { Val = "4BACC6" };

            accent5Color1.Append(rgbColorModelHex7);

            var accent6Color1 = new Accent6Color();
            var rgbColorModelHex8 = new RgbColorModelHex { Val = "F79646" };

            accent6Color1.Append(rgbColorModelHex8);

            var hyperlink1 = new Hyperlink();
            var rgbColorModelHex9 = new RgbColorModelHex { Val = "0000FF" };

            hyperlink1.Append(rgbColorModelHex9);

            var followedHyperlinkColor1 = new FollowedHyperlinkColor();
            var rgbColorModelHex10 = new RgbColorModelHex { Val = "800080" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            var fontScheme2 = new FontScheme { Name = "Office" };

            var majorFont1 = new MajorFont();
            var latinFont1 = new LatinFont { Typeface = "Cambria" };
            var eastAsianFont1 = new EastAsianFont { Typeface = "" };
            var complexScriptFont1 = new ComplexScriptFont { Typeface = "" };
            var supplementalFont1 = new SupplementalFont { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            var supplementalFont2 = new SupplementalFont { Script = "Hang", Typeface = "맑은 고딕" };
            var supplementalFont3 = new SupplementalFont { Script = "Hans", Typeface = "宋体" };
            var supplementalFont4 = new SupplementalFont { Script = "Hant", Typeface = "新細明體" };
            var supplementalFont5 = new SupplementalFont { Script = "Arab", Typeface = "Times New Roman" };
            var supplementalFont6 = new SupplementalFont { Script = "Hebr", Typeface = "Times New Roman" };
            var supplementalFont7 = new SupplementalFont { Script = "Thai", Typeface = "Tahoma" };
            var supplementalFont8 = new SupplementalFont { Script = "Ethi", Typeface = "Nyala" };
            var supplementalFont9 = new SupplementalFont { Script = "Beng", Typeface = "Vrinda" };
            var supplementalFont10 = new SupplementalFont { Script = "Gujr", Typeface = "Shruti" };
            var supplementalFont11 = new SupplementalFont { Script = "Khmr", Typeface = "MoolBoran" };
            var supplementalFont12 = new SupplementalFont { Script = "Knda", Typeface = "Tunga" };
            var supplementalFont13 = new SupplementalFont { Script = "Guru", Typeface = "Raavi" };
            var supplementalFont14 = new SupplementalFont { Script = "Cans", Typeface = "Euphemia" };
            var supplementalFont15 = new SupplementalFont { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            var supplementalFont16 = new SupplementalFont { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            var supplementalFont17 = new SupplementalFont { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            var supplementalFont18 = new SupplementalFont { Script = "Thaa", Typeface = "MV Boli" };
            var supplementalFont19 = new SupplementalFont { Script = "Deva", Typeface = "Mangal" };
            var supplementalFont20 = new SupplementalFont { Script = "Telu", Typeface = "Gautami" };
            var supplementalFont21 = new SupplementalFont { Script = "Taml", Typeface = "Latha" };
            var supplementalFont22 = new SupplementalFont { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            var supplementalFont23 = new SupplementalFont { Script = "Orya", Typeface = "Kalinga" };
            var supplementalFont24 = new SupplementalFont { Script = "Mlym", Typeface = "Kartika" };
            var supplementalFont25 = new SupplementalFont { Script = "Laoo", Typeface = "DokChampa" };
            var supplementalFont26 = new SupplementalFont { Script = "Sinh", Typeface = "Iskoola Pota" };
            var supplementalFont27 = new SupplementalFont { Script = "Mong", Typeface = "Mongolian Baiti" };
            var supplementalFont28 = new SupplementalFont { Script = "Viet", Typeface = "Times New Roman" };
            var supplementalFont29 = new SupplementalFont { Script = "Uigh", Typeface = "Microsoft Uighur" };
            var supplementalFont30 = new SupplementalFont { Script = "Geor", Typeface = "Sylfaen" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);
            majorFont1.Append(supplementalFont30);

            var minorFont1 = new MinorFont();
            var latinFont2 = new LatinFont { Typeface = "Calibri" };
            var eastAsianFont2 = new EastAsianFont { Typeface = "" };
            var complexScriptFont2 = new ComplexScriptFont { Typeface = "" };
            var supplementalFont31 = new SupplementalFont { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            var supplementalFont32 = new SupplementalFont { Script = "Hang", Typeface = "맑은 고딕" };
            var supplementalFont33 = new SupplementalFont { Script = "Hans", Typeface = "宋体" };
            var supplementalFont34 = new SupplementalFont { Script = "Hant", Typeface = "新細明體" };
            var supplementalFont35 = new SupplementalFont { Script = "Arab", Typeface = "Arial" };
            var supplementalFont36 = new SupplementalFont { Script = "Hebr", Typeface = "Arial" };
            var supplementalFont37 = new SupplementalFont { Script = "Thai", Typeface = "Tahoma" };
            var supplementalFont38 = new SupplementalFont { Script = "Ethi", Typeface = "Nyala" };
            var supplementalFont39 = new SupplementalFont { Script = "Beng", Typeface = "Vrinda" };
            var supplementalFont40 = new SupplementalFont { Script = "Gujr", Typeface = "Shruti" };
            var supplementalFont41 = new SupplementalFont { Script = "Khmr", Typeface = "DaunPenh" };
            var supplementalFont42 = new SupplementalFont { Script = "Knda", Typeface = "Tunga" };
            var supplementalFont43 = new SupplementalFont { Script = "Guru", Typeface = "Raavi" };
            var supplementalFont44 = new SupplementalFont { Script = "Cans", Typeface = "Euphemia" };
            var supplementalFont45 = new SupplementalFont { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            var supplementalFont46 = new SupplementalFont { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            var supplementalFont47 = new SupplementalFont { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            var supplementalFont48 = new SupplementalFont { Script = "Thaa", Typeface = "MV Boli" };
            var supplementalFont49 = new SupplementalFont { Script = "Deva", Typeface = "Mangal" };
            var supplementalFont50 = new SupplementalFont { Script = "Telu", Typeface = "Gautami" };
            var supplementalFont51 = new SupplementalFont { Script = "Taml", Typeface = "Latha" };
            var supplementalFont52 = new SupplementalFont { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            var supplementalFont53 = new SupplementalFont { Script = "Orya", Typeface = "Kalinga" };
            var supplementalFont54 = new SupplementalFont { Script = "Mlym", Typeface = "Kartika" };
            var supplementalFont55 = new SupplementalFont { Script = "Laoo", Typeface = "DokChampa" };
            var supplementalFont56 = new SupplementalFont { Script = "Sinh", Typeface = "Iskoola Pota" };
            var supplementalFont57 = new SupplementalFont { Script = "Mong", Typeface = "Mongolian Baiti" };
            var supplementalFont58 = new SupplementalFont { Script = "Viet", Typeface = "Arial" };
            var supplementalFont59 = new SupplementalFont { Script = "Uigh", Typeface = "Microsoft Uighur" };
            var supplementalFont60 = new SupplementalFont { Script = "Geor", Typeface = "Sylfaen" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);
            minorFont1.Append(supplementalFont59);
            minorFont1.Append(supplementalFont60);

            fontScheme2.Append(majorFont1);
            fontScheme2.Append(minorFont1);

            var formatScheme1 = new FormatScheme { Name = "Office" };

            var fillStyleList1 = new FillStyleList();

            var solidFill1 = new SolidFill();
            var schemeColor1 = new SchemeColor { Val = SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            var gradientFill1 = new GradientFill { RotateWithShape = true };

            var gradientStopList1 = new GradientStopList();

            var gradientStop1 = new GradientStop { Position = 0 };

            var schemeColor2 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var tint1 = new Tint { Val = 50000 };
            var saturationModulation1 = new SaturationModulation { Val = 300000 };

            schemeColor2.Append(tint1);
            schemeColor2.Append(saturationModulation1);

            gradientStop1.Append(schemeColor2);

            var gradientStop2 = new GradientStop { Position = 35000 };

            var schemeColor3 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var tint2 = new Tint { Val = 37000 };
            var saturationModulation2 = new SaturationModulation { Val = 300000 };

            schemeColor3.Append(tint2);
            schemeColor3.Append(saturationModulation2);

            gradientStop2.Append(schemeColor3);

            var gradientStop3 = new GradientStop { Position = 100000 };

            var schemeColor4 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var tint3 = new Tint { Val = 15000 };
            var saturationModulation3 = new SaturationModulation { Val = 350000 };

            schemeColor4.Append(tint3);
            schemeColor4.Append(saturationModulation3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            var linearGradientFill1 = new LinearGradientFill { Angle = 16200000, Scaled = true };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            var gradientFill2 = new GradientFill { RotateWithShape = true };

            var gradientStopList2 = new GradientStopList();

            var gradientStop4 = new GradientStop { Position = 0 };

            var schemeColor5 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var shade1 = new Shade { Val = 51000 };
            var saturationModulation4 = new SaturationModulation { Val = 130000 };

            schemeColor5.Append(shade1);
            schemeColor5.Append(saturationModulation4);

            gradientStop4.Append(schemeColor5);

            var gradientStop5 = new GradientStop { Position = 80000 };

            var schemeColor6 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var shade2 = new Shade { Val = 93000 };
            var saturationModulation5 = new SaturationModulation { Val = 130000 };

            schemeColor6.Append(shade2);
            schemeColor6.Append(saturationModulation5);

            gradientStop5.Append(schemeColor6);

            var gradientStop6 = new GradientStop { Position = 100000 };

            var schemeColor7 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var shade3 = new Shade { Val = 94000 };
            var saturationModulation6 = new SaturationModulation { Val = 135000 };

            schemeColor7.Append(shade3);
            schemeColor7.Append(saturationModulation6);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            var linearGradientFill2 = new LinearGradientFill { Angle = 16200000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            var lineStyleList1 = new LineStyleList();

            var outline1 = new Outline { Width = 9525, CapType = LineCapValues.Flat, CompoundLineType = CompoundLineValues.Single, Alignment = PenAlignmentValues.Center };

            var solidFill2 = new SolidFill();

            var schemeColor8 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var shade4 = new Shade { Val = 95000 };
            var saturationModulation7 = new SaturationModulation { Val = 105000 };

            schemeColor8.Append(shade4);
            schemeColor8.Append(saturationModulation7);

            solidFill2.Append(schemeColor8);
            var presetDash1 = new PresetDash { Val = PresetLineDashValues.Solid };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);

            var outline2 = new Outline { Width = 25400, CapType = LineCapValues.Flat, CompoundLineType = CompoundLineValues.Single, Alignment = PenAlignmentValues.Center };

            var solidFill3 = new SolidFill();
            var schemeColor9 = new SchemeColor { Val = SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            var presetDash2 = new PresetDash { Val = PresetLineDashValues.Solid };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);

            var outline3 = new Outline { Width = 38100, CapType = LineCapValues.Flat, CompoundLineType = CompoundLineValues.Single, Alignment = PenAlignmentValues.Center };

            var solidFill4 = new SolidFill();
            var schemeColor10 = new SchemeColor { Val = SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            var presetDash3 = new PresetDash { Val = PresetLineDashValues.Solid };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            var effectStyleList1 = new EffectStyleList();

            var effectStyle1 = new EffectStyle();

            var effectList1 = new EffectList();

            var outerShadow1 = new OuterShadow { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false };

            var rgbColorModelHex11 = new RgbColorModelHex { Val = "000000" };
            var alpha1 = new Alpha { Val = 38000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList1.Append(outerShadow1);

            effectStyle1.Append(effectList1);

            var effectStyle2 = new EffectStyle();

            var effectList2 = new EffectList();

            var outerShadow2 = new OuterShadow { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            var rgbColorModelHex12 = new RgbColorModelHex { Val = "000000" };
            var alpha2 = new Alpha { Val = 35000 };

            rgbColorModelHex12.Append(alpha2);

            outerShadow2.Append(rgbColorModelHex12);

            effectList2.Append(outerShadow2);

            effectStyle2.Append(effectList2);

            var effectStyle3 = new EffectStyle();

            var effectList3 = new EffectList();

            var outerShadow3 = new OuterShadow { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            var rgbColorModelHex13 = new RgbColorModelHex { Val = "000000" };
            var alpha3 = new Alpha { Val = 35000 };

            rgbColorModelHex13.Append(alpha3);

            outerShadow3.Append(rgbColorModelHex13);

            effectList3.Append(outerShadow3);

            var scene3DType1 = new Scene3DType();

            var camera1 = new Camera { Preset = PresetCameraValues.OrthographicFront };
            var rotation1 = new Rotation { Latitude = 0, Longitude = 0, Revolution = 0 };

            camera1.Append(rotation1);

            var lightRig1 = new LightRig { Rig = LightRigValues.ThreePoints, Direction = LightRigDirectionValues.Top };
            var rotation2 = new Rotation { Latitude = 0, Longitude = 0, Revolution = 1200000 };

            lightRig1.Append(rotation2);

            scene3DType1.Append(camera1);
            scene3DType1.Append(lightRig1);

            var shape3DType1 = new Shape3DType();
            var bevelTop1 = new BevelTop { Width = 63500L, Height = 25400L };

            shape3DType1.Append(bevelTop1);

            effectStyle3.Append(effectList3);
            effectStyle3.Append(scene3DType1);
            effectStyle3.Append(shape3DType1);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            var backgroundFillStyleList1 = new BackgroundFillStyleList();

            var solidFill5 = new SolidFill();
            var schemeColor11 = new SchemeColor { Val = SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            var gradientFill3 = new GradientFill { RotateWithShape = true };

            var gradientStopList3 = new GradientStopList();

            var gradientStop7 = new GradientStop { Position = 0 };

            var schemeColor12 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var tint4 = new Tint { Val = 40000 };
            var saturationModulation8 = new SaturationModulation { Val = 350000 };

            schemeColor12.Append(tint4);
            schemeColor12.Append(saturationModulation8);

            gradientStop7.Append(schemeColor12);

            var gradientStop8 = new GradientStop { Position = 40000 };

            var schemeColor13 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var tint5 = new Tint { Val = 45000 };
            var shade5 = new Shade { Val = 99000 };
            var saturationModulation9 = new SaturationModulation { Val = 350000 };

            schemeColor13.Append(tint5);
            schemeColor13.Append(shade5);
            schemeColor13.Append(saturationModulation9);

            gradientStop8.Append(schemeColor13);

            var gradientStop9 = new GradientStop { Position = 100000 };

            var schemeColor14 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var shade6 = new Shade { Val = 20000 };
            var saturationModulation10 = new SaturationModulation { Val = 255000 };

            schemeColor14.Append(shade6);
            schemeColor14.Append(saturationModulation10);

            gradientStop9.Append(schemeColor14);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);

            var pathGradientFill1 = new PathGradientFill { Path = PathShadeValues.Circle };
            var fillToRectangle1 = new FillToRectangle { Left = 50000, Top = -80000, Right = 50000, Bottom = 180000 };

            pathGradientFill1.Append(fillToRectangle1);

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(pathGradientFill1);

            var gradientFill4 = new GradientFill { RotateWithShape = true };

            var gradientStopList4 = new GradientStopList();

            var gradientStop10 = new GradientStop { Position = 0 };

            var schemeColor15 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var tint6 = new Tint { Val = 80000 };
            var saturationModulation11 = new SaturationModulation { Val = 300000 };

            schemeColor15.Append(tint6);
            schemeColor15.Append(saturationModulation11);

            gradientStop10.Append(schemeColor15);

            var gradientStop11 = new GradientStop { Position = 100000 };

            var schemeColor16 = new SchemeColor { Val = SchemeColorValues.PhColor };
            var shade7 = new Shade { Val = 30000 };
            var saturationModulation12 = new SaturationModulation { Val = 200000 };

            schemeColor16.Append(shade7);
            schemeColor16.Append(saturationModulation12);

            gradientStop11.Append(schemeColor16);

            gradientStopList4.Append(gradientStop10);
            gradientStopList4.Append(gradientStop11);

            var pathGradientFill2 = new PathGradientFill { Path = PathShadeValues.Circle };
            var fillToRectangle2 = new FillToRectangle { Left = 50000, Top = 50000, Right = 50000, Bottom = 50000 };

            pathGradientFill2.Append(fillToRectangle2);

            gradientFill4.Append(gradientStopList4);
            gradientFill4.Append(pathGradientFill2);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(gradientFill3);
            backgroundFillStyleList1.Append(gradientFill4);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme2);
            themeElements1.Append(formatScheme1);
            var objectDefaults1 = new ObjectDefaults();
            var extraColorSchemeList1 = new ExtraColorSchemeList();

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);

            return theme1;
        }

    }
}
