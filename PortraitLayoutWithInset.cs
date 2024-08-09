using ArcGIS.Core.CIM;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Portal;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Internal.Mapping.Operations;
using ArcGIS.Desktop.Layouts;
using ArcGIS.Desktop.Mapping;
using System;
using System.Collections.Generic;
using System.Linq;
using static System.Net.WebRequestMethods;

namespace AutomatedLayoutProduction
{
    internal class PortraitLayoutWithInset : Button
    {
        protected override async void OnClick()
        {
            // Create layout with landscape orientation on the worker thread
            Layout lyt = await QueuedTask.Run(() =>
            {
                var newLayout = LayoutFactory.Instance.CreateLayout(8.5, 11, ArcGIS.Core.Geometry.LinearUnit.Inches);
                newLayout.SetName("Portrait Layout with Inset");
                StyleProjectItem stylePrjItm = Project.Current.GetItems<StyleProjectItem>().FirstOrDefault(item => item.Name == "ArcGIS 2D");
                // Build geometry for the map frame to take up most of the layout
                Coordinate2D ll = new(0.125, 0.125);
                Coordinate2D ur = new(8.375, 10.2554);
                ArcGIS.Core.Geometry.Envelope env = EnvelopeBuilderEx.CreateEnvelope(ll, ur);

                // Reference the active map
                Map map = MapView.Active.Map;

                // Create a map frame element and add it to the layout
                MapFrame mfElm = ElementFactory.Instance.CreateMapFrameElement(newLayout, env, map);
                mfElm.SetName("Core Map Frame");

                // Calculate the combined extent of all active layers
                if (map != null)
                {
                    ArcGIS.Core.Geometry.Envelope combinedExtent = null;
                    foreach (var layer in map.GetLayersAsFlattenedList().OfType<FeatureLayer>())
                    {
                        if (layer.IsVisible)
                        {
                            var layerExtent = layer.QueryExtent();
                            if (combinedExtent == null)
                            {
                                combinedExtent = layerExtent;
                            }
                            else
                            {
                                combinedExtent = combinedExtent.Union(layerExtent);
                            }
                        }
                    }

                    // Set the map frame extent to the combined extent of all active layers
                    if (combinedExtent != null)
                    {
                        mfElm.SetCamera(combinedExtent);

                        // Zoom out 25 percent
                        Camera cam = mfElm.Camera;
                        cam.Scale *= 1.25;
                        mfElm.SetCamera(cam);
                    }
                }

                //Create Inset Mapframe
                Coordinate2D smll = new(6, 8.375);
                Coordinate2D smur = new(8.375, 10.875);
                ArcGIS.Core.Geometry.Envelope env2 = EnvelopeBuilderEx.CreateEnvelope(smll, smur);

                Map map2 = MapFactory.Instance.CopyMap(map);

                QueuedTask.Run(() =>
                {
                    IReadOnlyList<Layer> layers = map2.Layers.ToList();
                    map2.RemoveLayers(layers);
                });

                map2.SetName("Inset Map: Low Zoom");

                var basemap = LayerFactory.Instance.CreateLayer(new Uri("https://www.arcgis.com/sharing/rest/content/items/5e9b3685f4c24d8781073dd928ebda50/resources/styles/root.json"), map2);
                basemap.SetName("Dark Grey Base");

                var citiesParam = new FeatureLayerCreationParams(new Uri(@"https://services.arcgis.com/P3ePLMYs2RVChkJx/ArcGIS/rest/services/USA_Major_Cities_/FeatureServer/0"))
                {
                    Name = "Major Cities",
                    DefinitionQuery = new DefinitionQuery(whereClause: "POP_CLASS >= 8", name: "Population Definition"),
                    RendererDefinition = new SimpleRendererDefinition()
                    {
                        SymbolTemplate = SymbolFactory.Instance.ConstructPointSymbol(CIMColor.CreateRGBColor(255, 144, 200, 100), 9, SimpleMarkerStyle.Diamond).MakeSymbolReference(),
                    },


                };

                var cities = LayerFactory.Instance.CreateLayer<FeatureLayer>(citiesParam, map2);
                var citiesDef = cities.GetDefinition() as CIMFeatureLayer;
                var citiesLabel = citiesDef.LabelClasses.FirstOrDefault();
                if (citiesLabel != null)
                {
                    var citiesText = SymbolFactory.Instance.ConstructTextSymbol(CIMColor.CreateRGBColor(177, 48, 177, 100), 8, "Arial", "Regular");
                    citiesText.HaloSize = .2;
                    citiesText.HaloSymbol = SymbolFactory.Instance.ConstructPolygonSymbol(CIMColor.CreateRGBColor(255, 255, 255, 75), SimpleFillStyle.Solid);
                    citiesLabel.TextSymbol.Symbol = citiesText;
                    citiesLabel.MaplexLabelPlacementProperties.FeatureType = LabelFeatureType.Point;
                    citiesLabel.MaplexLabelPlacementProperties.NeverRemoveLabel = true;
                    citiesLabel.MaplexLabelPlacementProperties.IsOffsetFromFeatureGeometry = true;
                    citiesDef.LabelClasses = new[] { citiesLabel };
                    cities.SetDefinition(citiesDef);
                    cities.SetLabelVisibility(true);
                };

                SymbolStyleItem point = stylePrjItm.SearchSymbols(StyleItemType.PointSymbol, "Esri Pin 1")[0];
                CIMPointSymbol pointCIM = point.Symbol as CIMPointSymbol;
                pointCIM.SetSize(20);

                CIMSymbolReference pointCIMRef = new()
                {
                    Symbol = pointCIM,
                };

                CIMStroke extentOutline = SymbolFactory.Instance.ConstructStroke(ColorFactory.Instance.RedRGB, 1.0, SimpleLineStyle.Solid);
                CIMPolygonSymbol extentPoly = SymbolFactory.Instance.ConstructPolygonSymbol(ColorFactory.Instance.BlackRGB, SimpleFillStyle.Null, extentOutline);

                CIMSymbolReference extentPolyRef = new()
                {
                    Symbol = extentPoly,
                };

                MapFrame M2 = ElementFactory.Instance.CreateMapFrameElement(newLayout, env2, map2, "Inset Map Frame");

                var m2CIM = M2.GetDefinition() as CIMMapFrame;
                var cIMExtentIndicator = new CIMExtentIndicator()
                {
                    SourceMapFrame = "Core Map Frame",
                    CollapseSize = 999999999999,
                    PointSymbol = pointCIMRef,
                    Symbol = extentPolyRef,
                };

                m2CIM.ExtentIndicators = [cIMExtentIndicator];
                m2CIM.ExtentIndicators[0] = cIMExtentIndicator;
                M2.SetDefinition(m2CIM);

                // Add the portal item as a new layer
                CIMStroke outline = SymbolFactory.Instance.ConstructStroke(CIMColor.CreateRGBColor(50, 75, 33, 100),2,SimpleLineStyle.Solid);
                CIMFill innard = SymbolFactory.Instance.ConstructSolidFill(CIMColor.CreateRGBColor(0, 0, 0, 0));
                CIMStroke haloOutline = SymbolFactory.Instance.ConstructStroke(CIMColor.CreateRGBColor(0, 0, 0, 0),0,SimpleLineStyle.Solid);
                string portalItemID = "6b3112d1c2264cd39cfa1d109fa73283";
                Item portalItem = ItemFactory.Instance.Create(portalItemID, ItemFactory.ItemType.PortalItem);
                var layerParams = new FeatureLayerCreationParams(portalItem as PortalItem);
                layerParams.RendererDefinition = new SimpleRendererDefinition()
                {
                    SymbolTemplate = SymbolFactory.Instance.ConstructPolygonSymbol(innard, outline).MakeSymbolReference(),
                    
                };
                var featureLayer = LayerFactory.Instance.CreateLayer<FeatureLayer>(layerParams, map2);
                featureLayer.SetLabelVisibility(true);
                var featureLayerDef = featureLayer.GetDefinition() as CIMFeatureLayer;
                var labelClass = featureLayerDef.LabelClasses.FirstOrDefault();
                if (labelClass != null)
                {
                    var texSymbol = SymbolFactory.Instance.ConstructTextSymbol(ColorFactory.Instance.CreateRGBColor(0, 122, 194, 100), 12, "Arial", "Regular");
                    texSymbol.HaloSize = .2;
                    texSymbol.HaloSymbol = SymbolFactory.Instance.ConstructPolygonSymbol(CIMColor.CreateRGBColor(255, 255, 255, 75),SimpleFillStyle.Solid, haloOutline);
                    labelClass.TextSymbol.Symbol = texSymbol;
                    labelClass.MaplexLabelPlacementProperties.FeatureType = LabelFeatureType.Polygon;
                    labelClass.MaplexLabelPlacementProperties.PolygonPlacementMethod = MaplexPolygonPlacementMethod.RepeatAlongBoundary;
                    labelClass.MaplexLabelPlacementProperties.BoundaryLabelingAllowSingleSided = false;
                    labelClass.MaplexLabelPlacementProperties.BoundaryLabelingSingleSidedOnLine = false;
                    labelClass.MaplexLabelPlacementProperties.LabelBuffer = 50;
                    labelClass.MaplexLabelPlacementProperties.IsLabelBufferHardConstraint = true;
                    labelClass.MaplexLabelPlacementProperties.CanPlaceLabelOutsidePolygon = false;
                    labelClass.MaplexLabelPlacementProperties.AvoidPolygonHoles = true;
                    labelClass.MaplexLabelPlacementProperties.RemoveAmbiguousLabels = MaplexRemoveAmbiguousLabelsType.All;
                    labelClass.MaplexLabelPlacementProperties.PolygonBoundaryWeight = 0;
                    labelClass.MaplexLabelPlacementProperties.FeatureWeight = 1000;
                    labelClass.MaplexLabelPlacementProperties.PreferHorizontalPlacement = true;
                    labelClass.MaplexLabelPlacementProperties.PrimaryOffset = 5;
                    labelClass.MaplexLabelPlacementProperties.IsOffsetFromFeatureGeometry = true;
                    labelClass.MaplexLabelPlacementProperties.AlignLabelToLineDirection = false;
                    labelClass.MaplexLabelPlacementProperties.RepeatLabel = false;
                    featureLayerDef.LabelClasses = new[] { labelClass };
                    featureLayerDef.Transparency = 0;
                    featureLayer.SetDefinition(featureLayerDef);
                };

                //extent for inset map frame
                Camera cam2 = M2.Camera;
                cam2 = mfElm.Camera;
                cam2.Scale *= 5;
                M2.SetCamera(cam2);

                // Add Title bar graphic element
                Coordinate2D tBar_ll = new(0.125, 10.25);
                Coordinate2D tBar_ur = new(6, 10.875);
                ArcGIS.Core.Geometry.Envelope tBarEnv = EnvelopeBuilderEx.CreateEnvelope(tBar_ll, tBar_ur);

                CIMStroke tBarStroke = SymbolFactory.Instance.ConstructStroke(ColorFactory.Instance.BlackRGB);
                CIMPolygonSymbol tBarPolySym = SymbolFactory.Instance.ConstructPolygonSymbol(ColorFactory.Instance.CreateRGBColor(0, 122, 194, 30));
                CIMGraphic tBarGraphic = GraphicFactory.Instance.CreateSimpleGraphic(tBarEnv, tBarPolySym);

                var tBarElInfo = new ElementInfo()
                {
                    Anchor = Anchor.BottomLeftCorner,
                };

                var tBar = ElementFactory.Instance.CreateGraphicElement(newLayout, tBarGraphic, "Title Bar", true, tBarElInfo);


                // Add the map name as a title in the bottom right corner of the map Frame
                string title = $@"<dyn type=""mapFrame"" name=""Core Map Frame"" property=""mapName""/>";
                Coordinate2D titlePosition = new(6.25, 1.3);

                SymbolStyleItem polyHalo = stylePrjItm.SearchSymbols(StyleItemType.PolygonSymbol, "Glacier")[0];

                CIMPolygonSymbol glacierHalo = polyHalo.Symbol as CIMPolygonSymbol;

                CIMTextSymbol cimTitle = SymbolFactory.Instance.ConstructTextSymbol(ColorFactory.Instance.BlackRGB, 24, "Arial", "Bold");
                cimTitle.HaloSize = 1;
                cimTitle.HaloSymbol = glacierHalo;

                var titleInfo = new ElementInfo()
                {
                    Anchor = Anchor.BottomMidPoint,

                };

                var titleText = ElementFactory.Instance.CreateTextGraphicElement(newLayout, TextType.PointText, titlePosition.ToMapPoint(), cimTitle, title, "Core Map Title", true, titleInfo);

                var titleAnchor = titleText.GetAnchor();
                {
                    titleAnchor = Anchor.BottomMidPoint;

                    titleText.SetAnchor(titleAnchor);
                };

                var titleAnchorPoint = titleText.GetAnchorPoint();
                {
                    titleAnchorPoint = titlePosition;

                    titleText.SetAnchorPoint(titleAnchorPoint);
                };

                //Add title to second map frame
                string title2 = $@"<dyn type=""mapFrame""name=""Inset Map Frame"" property=""mapName""/>";
                Coordinate2D title2Position = new(8.25, 8.638);

                CIMTextSymbol cimTitle2 = SymbolFactory.Instance.ConstructTextSymbol(ColorFactory.Instance.BlackRGB, 10, "Arial", "Bold");
                cimTitle2.HaloSize = .5;
                cimTitle2.HaloSymbol = glacierHalo;

                var titleInfo2 = new ElementInfo()
                {
                    Anchor = Anchor.TopRightCorner,
                };

                var titleText2 = ElementFactory.Instance.CreateTextGraphicElement(newLayout, TextType.PointText, title2Position.ToMapPoint(), cimTitle2, title2, "Inset Map Title", true, titleInfo2);

                var title2Anchor = titleText2.GetAnchor();
                {
                    title2Anchor = Anchor.TopRightCorner;

                    titleText2.SetAnchor(title2Anchor);
                }

                var title2AnchorPoint = titleText2.GetAnchorPoint();
                {
                    title2AnchorPoint = title2Position;

                    titleText2.SetAnchorPoint(title2AnchorPoint);
                }

                // add project title as main title
                string mainTitle = $@"<dyn type=""project"" property=""name""/>";
                Coordinate2D mainTitlePosition = new(0.25, 10.2693);

                CIMTextSymbol cimMainTitle = SymbolFactory.Instance.ConstructTextSymbol(ColorFactory.Instance.BlackRGB, 36, "Arial", "Bold");
                cimMainTitle.HaloSize = 1;
                cimMainTitle.HaloSymbol = glacierHalo;

                var mainTitleInfo = new ElementInfo()
                {
                    Anchor = Anchor.BottomLeftCorner,
                };

                var mainTitleText = ElementFactory.Instance.CreateTextGraphicElement(newLayout, TextType.PointText, mainTitlePosition.ToMapPoint(), cimMainTitle, mainTitle, "Project Title", true, mainTitleInfo);


                // Add a north arrow in the top right corner of the layout
                NorthArrowStyleItem naStyleItm = stylePrjItm.SearchNorthArrows("ArcGIS North 10")[0];

                // Position the north arrow in the top right corner of the layout
                Coordinate2D naPosition = new(0.5421, 9.6308);

                var naInfo = new NorthArrowInfo()
                {
                    MapFrameName = mfElm.Name,
                    NorthArrowStyleItem = naStyleItm
                };

                var arrowElm = ElementFactory.Instance.CreateMapSurroundElement(
                    newLayout, naPosition.ToMapPoint(), naInfo, "North Arrow") as NorthArrow;
                arrowElm.SetHeight(.8);  // Adjust height as needed

                var northArrowCim = arrowElm.GetDefinition() as CIMMarkerNorthArrow;
                var naPolyFille = SymbolFactory.Instance.ConstructSolidFill(ColorFactory.Instance.CreateRGBColor(249, 249, 245, 100));
                var naPolyStroke = SymbolFactory.Instance.ConstructStroke(ColorFactory.Instance.WhiteRGB, 0);
                var halopoly = SymbolFactory.Instance.ConstructPolygonSymbol(naPolyFille, naPolyStroke);

                ((CIMPointSymbol)northArrowCim.PointSymbol.Symbol).HaloSymbol = halopoly;
                ((CIMPointSymbol)northArrowCim.PointSymbol.Symbol).HaloSize = 1;

                arrowElm.SetDefinition(northArrowCim);

                // Add a scale bar in the bottom left corner of the layout
                ScaleBarStyleItem sbStyleItm = stylePrjItm.SearchScaleBars("Double Alternating Scale Bar 1")[0];

                Coordinate2D sb_ll = new(.25, 2.125);
                Coordinate2D sb_ur = new(3.1082, 2.496);
                ArcGIS.Core.Geometry.Envelope sbenv = EnvelopeBuilderEx.CreateEnvelope(sb_ll, sb_ur);

                var sbInfo = new ScaleBarInfo()
                {
                    MapFrameName = mfElm.Name,
                    ScaleBarStyleItem = sbStyleItm,

                };

                var sbElm = ElementFactory.Instance.CreateMapSurroundElement(
                    newLayout, sbenv, sbInfo, "Scale Bar") as ScaleBar;
                sbElm.SetWidth(3.0);  // Adjust width as needed

                if (sbElm.GetDefinition() is CIMScaleBar cimScaleBar)
                {
                    cimScaleBar.Divisions = 4;
                    cimScaleBar.Subdivisions = 0;
                    cimScaleBar.LabelFrequency = ScaleBarFrequency.Divisions;
                    cimScaleBar.FittingStrategy = ScaleBarFittingStrategy.AdjustFrame;
                    sbElm.SetDefinition(cimScaleBar);
                }

                //Add rectangle around legend
                Coordinate2D legrec_ll = new(.125, .125);
                Coordinate2D legrec_ur = new(4, 2);
                ArcGIS.Core.Geometry.Envelope legrecEnv = EnvelopeBuilderEx.CreateEnvelope(legrec_ll, legrec_ur);

                CIMStroke legRecStroke = SymbolFactory.Instance.ConstructStroke(ColorFactory.Instance.BlackRGB);
                CIMPolygonSymbol legrecPolySym = SymbolFactory.Instance.ConstructPolygonSymbol(ColorFactory.Instance.CreateRGBColor(255, 255, 255, 50), SimpleFillStyle.Solid, legRecStroke);
                CIMGraphic rect = GraphicFactory.Instance.CreateSimpleGraphic(legrecEnv, legrecPolySym);

                var legRectInfo = new ElementInfo()
                {
                    Anchor = Anchor.BottomLeftCorner,
                };

                var legendRectangle = ElementFactory.Instance.CreateGraphicElement(newLayout, rect, "Legend Rectangle", true, legRectInfo);


                // Add a legend in the bottom left corner of the layout
                Coordinate2D leg_ul = new(0.25, 1.875);

                SymbolStyleItem textHalo = stylePrjItm.SearchSymbols(StyleItemType.PolygonSymbol, "Water Intermittent")[0];

                CIMPolygonSymbol waterIntermit = textHalo.Symbol as CIMPolygonSymbol;

                var legendInfo = new LegendInfo()
                {
                    MapFrameName = mfElm.Name,
                };

                var legendElm = ElementFactory.Instance.CreateMapSurroundElement(
                    newLayout, leg_ul.ToMapPoint(), legendInfo, "Legend") as Legend;

                var legAnchor = legendElm.GetAnchor();
                {
                    legAnchor = Anchor.TopLeftCorner;

                    legendElm.SetAnchor(legAnchor);
                }

                var legAnchorPoint = legendElm.GetAnchorPoint();
                {
                    legAnchorPoint = leg_ul;

                    legendElm.SetAnchorPoint(legAnchorPoint);
                }


                CIMTextSymbol tSym = SymbolFactory.Instance.ConstructTextSymbol(ColorFactory.Instance.BlackRGB, 12, "Arial", "Regular");
                tSym.HaloSymbol = waterIntermit;
                tSym.HaloSize = .25;

                CIMSymbolReference titleSym = new()
                {
                    Symbol = tSym,
                };

                if (legendElm.GetDefinition() is CIMLegend cimLegend)
                {
                    // Customize legend properties
                    cimLegend.ShowTitle = true;
                    cimLegend.Title = "Legend";
                    cimLegend.FittingStrategy = LegendFittingStrategy.AdjustFrame;
                    cimLegend.TitleSymbol = titleSym;

                    // Apply changes to the legend element
                    legendElm.SetDefinition(cimLegend);

                    // Modify each legend item directly
                    foreach (var item in cimLegend.Items)
                    {
                        item.ShowHeading = false;
                        item.ShowGroupLayerName = false;
                        item.ShowLayerName = false;
                        item.LabelSymbol = titleSym;
                        item.LayerNameSymbol = titleSym;
                    }

                    // Apply changes to the legend element again
                    legendElm.SetDefinition(cimLegend);
                }

                //Add Map Frame Description automated text to bottom right of layout
                string description = $@"<dyn type =""mapFrame"" name =""Core Map Frame"" property =""description"" />";

                Coordinate2D description_ll = new(4.375, 0.375);
                Coordinate2D description_ur = new(8.125, 1.25);
                ArcGIS.Core.Geometry.Envelope descEnvPosition = EnvelopeBuilderEx.CreateEnvelope(description_ll, description_ur);

                CIMTextSymbol cimDesc = SymbolFactory.Instance.ConstructTextSymbol(ColorFactory.Instance.BlackRGB, 8, "Arial", "Regular");
                cimDesc.HorizontalAlignment = HorizontalAlignment.Center;
                cimDesc.OffsetY = -5;

                var descInfo = new ElementInfo()
                {
                    Anchor = Anchor.BottomLeftCorner,

                };


                var descText = ElementFactory.Instance.CreateTextGraphicElement(newLayout, TextType.RectangleParagraph, descEnvPosition.Extent, cimDesc, description, "Core Frame Description", true, descInfo);
                
                var polySymbol = SymbolFactory.Instance.ConstructPolygonSymbol(ColorFactory.Instance.CreateRGBColor(255, 255, 255, 50), SimpleFillStyle.Solid);

                CIMGraphic cimDescGra = descText.GetGraphic();
                CIMParagraphTextGraphic cimDescText = cimDescGra as CIMParagraphTextGraphic;
                cimDescText.Frame.BorderSymbol.Symbol.SetSize(1);
                cimDescText.Frame.BorderSymbol.Symbol.SetColor(ColorFactory.Instance.BlackRGB);
                cimDescText.Frame.BorderCornerRounding = 45;
                cimDescText.Frame.BackgroundSymbol = polySymbol.MakeSymbolReference();
                cimDescText.Frame.BackgroundCornerRounding = 45;
                descText.SetGraphic(cimDescGra);

                //Remove Service Layer
                string serviceLayer = $@"<dyn type =""layout"" name =""Portrait Layout with Inset"" property = ""serviceLayerCredits""/>";
                Coordinate2D serviceLayer_ll = new(0, 0);

                CIMTextSymbol cimServiceLayer = SymbolFactory.Instance.ConstructTextSymbol(ColorFactory.Instance.CreateRGBColor(255,255,255,0), 10);
                var servText = ElementFactory.Instance.CreateTextGraphicElement(newLayout, TextType.PointText, serviceLayer_ll.ToMapPoint(), cimServiceLayer, serviceLayer, "Invisible Service Layer");

                //Group similar elements

                // Return the updated layout
                return newLayout;
            });

            // Open new layout on the GUI thread
            await ProApp.Panes.CreateLayoutPaneAsync(lyt);
        }
    }
}
