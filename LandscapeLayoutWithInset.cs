using ArcGIS.Core.CIM;
using ArcGIS.Core.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Geoprocessing;
using ArcGIS.Desktop.Core.Portal;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Layouts;
using ArcGIS.Desktop.Mapping;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AutomatedLayoutProduction
{
    internal class LandscapeLayoutWithInset : Button
    {
        protected override async void OnClick()
        {
            // Create layout with landscape orientation on the worker thread
            Layout lyt = await QueuedTask.Run(() =>
            {
                var newLayout = LayoutFactory.Instance.CreateLayout(11, 8.5, ArcGIS.Core.Geometry.LinearUnit.Inches);
                newLayout.SetName("Landscape Layout with Inset");
                StyleProjectItem stylePrjItm = Project.Current.GetItems<StyleProjectItem>().FirstOrDefault(item => item.Name == "ArcGIS 2D");
                // Build geometry for the map frame to take up most of the layout
                Coordinate2D ll = new(0.125, 0.125);
                Coordinate2D ur = new(7, 7.75);
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

                        // Zoom out 20 percent
                        Camera cam = mfElm.Camera;
                        cam.Scale *= 1.20;
                        mfElm.SetCamera(cam);
                    }
                }

                //Create Inset Map Frame
                Coordinate2D smll = new(7, 4.125);
                Coordinate2D smur = new(10.875, 8.375);
                ArcGIS.Core.Geometry.Envelope env2 = EnvelopeBuilderEx.CreateEnvelope(smll, smur);

                Map map2 = MapFactory.Instance.CopyMap(map);

                QueuedTask.Run(() =>
                {
                    IReadOnlyList<Layer> layers = map2.Layers.ToList();
                    map2.RemoveLayers(layers);
                });

                var basemap = LayerFactory.Instance.CreateLayer(new Uri("https://www.arcgis.com/sharing/rest/content/items/5e9b3685f4c24d8781073dd928ebda50/resources/styles/root.json"), map2);
                var cities = LayerFactory.Instance.CreateLayer(new Uri("https://services.arcgis.com/P3ePLMYs2RVChkJx/ArcGIS/rest/services/USA_Major_Cities_/FeatureServer/0"), map2);
                map2.SetName("Inset Map: Low Zoom");

                SymbolStyleItem point = stylePrjItm.SearchSymbols(StyleItemType.PointSymbol, "Esri Pin 1")[0];
                CIMPointSymbol pointCIM = point.Symbol as CIMPointSymbol;
                pointCIM.SetSize(30);

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
                CIMStroke outline = SymbolFactory.Instance.ConstructStroke(CIMColor.CreateRGBColor(50, 75, 33, 100), 2, SimpleLineStyle.Solid);
                CIMFill innard = SymbolFactory.Instance.ConstructSolidFill(CIMColor.CreateRGBColor(0, 0, 0, 0));
                CIMStroke haloOutline = SymbolFactory.Instance.ConstructStroke(CIMColor.CreateRGBColor(0, 0, 0, 0), 0, SimpleLineStyle.Solid);
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
                    texSymbol.HaloSymbol = SymbolFactory.Instance.ConstructPolygonSymbol(CIMColor.CreateRGBColor(255, 255, 255, 75), SimpleFillStyle.Solid, haloOutline);
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
                cam2.Scale *= 4;
                M2.SetCamera(cam2);

                // Add Title bar graphic element
                Coordinate2D tBar_ll = new(0.125, 7.75);
                Coordinate2D tBar_ur = new(7, 8.375);
                ArcGIS.Core.Geometry.Envelope tBarEnv = EnvelopeBuilderEx.CreateEnvelope(tBar_ll, tBar_ur);

                CIMStroke tBarStroke = SymbolFactory.Instance.ConstructStroke(ColorFactory.Instance.BlackRGB);
                CIMPolygonSymbol tBarPolySym = SymbolFactory.Instance.ConstructPolygonSymbol(ColorFactory.Instance.CreateRGBColor(0, 122, 194, 30));
                CIMGraphic tBarGraphic = GraphicFactory.Instance.CreateSimpleGraphic(tBarEnv, tBarPolySym);

                var tBarElInfo = new ElementInfo()
                {
                    Anchor = Anchor.BottomLeftCorner,
                };

                var tBar = ElementFactory.Instance.CreateGraphicElement(newLayout, tBarGraphic, "Title Bar", true, tBarElInfo);


                // Add the map name as a title in the top left corner of the map Frame
                string title = $@"<dyn type=""mapFrame"" name=""Core Map Frame"" property=""mapName""/>";
                Coordinate2D titlePosition = new(0.25, 7.2248);

                SymbolStyleItem polyHalo = stylePrjItm.SearchSymbols(StyleItemType.PolygonSymbol, "Glacier")[0];

                CIMPolygonSymbol glacierHalo = polyHalo.Symbol as CIMPolygonSymbol;

                CIMTextSymbol cimTitle = SymbolFactory.Instance.ConstructTextSymbol(ColorFactory.Instance.BlackRGB, 24, "Arial", "Bold");
                cimTitle.HaloSize = 1;
                cimTitle.HaloSymbol = glacierHalo;

                var titleInfo = new ElementInfo()
                {
                    Anchor = Anchor.BottomLeftCorner,

                };

                var titleText = ElementFactory.Instance.CreateTextGraphicElement(newLayout, TextType.PointText, titlePosition.ToMapPoint(), cimTitle, title, "Core Map Title", true, titleInfo);

                //Add title to second map frame
                string title2 = $@"<dyn type=""mapFrame""name=""Inset Map Frame"" property=""mapName""/>";
                Coordinate2D title2Position = new(10.75, 8.25);

                CIMTextSymbol cimTitle2 = SymbolFactory.Instance.ConstructTextSymbol(ColorFactory.Instance.BlackRGB, 16, "Arial", "Bold");
                cimTitle2.HaloSize = 1;
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
                Coordinate2D mainTitlePosition = new(0.1875, 7.75);

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
                Coordinate2D naPosition = new(6.5829, 7.125);

                var naInfo = new NorthArrowInfo()
                {
                    MapFrameName = mfElm.Name,
                    NorthArrowStyleItem = naStyleItm
                };

                var arrowElm = ElementFactory.Instance.CreateMapSurroundElement(
                    newLayout, naPosition.ToMapPoint(), naInfo, "North Arrow") as NorthArrow;
                arrowElm.SetHeight(.8);  // Adjust height as needed

                var northArrowCim = arrowElm.GetDefinition() as CIMMarkerNorthArrow;
                var naPolyFille = SymbolFactory.Instance.ConstructSolidFill(ColorFactory.Instance.CreateRGBColor(0, 174, 239, 100));
                var naPolyStroke = SymbolFactory.Instance.ConstructStroke(ColorFactory.Instance.WhiteRGB, 0);
                var halopoly = SymbolFactory.Instance.ConstructPolygonSymbol(naPolyFille, naPolyStroke);

                ((CIMPointSymbol)northArrowCim.PointSymbol.Symbol).HaloSymbol = halopoly;
                ((CIMPointSymbol)northArrowCim.PointSymbol.Symbol).HaloSize = 1;

                arrowElm.SetDefinition(northArrowCim);

                // Add a scale bar in the bottom left corner of the layout
                ScaleBarStyleItem sbStyleItm = stylePrjItm.SearchScaleBars("Double Alternating Scale Bar 1")[0];

                Coordinate2D sb_ll = new(2.1843, 0.375);
                Coordinate2D sb_ur = new(5.0657, 0.746);
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

                // Add a legend in the bottom right corner of the layout
                Coordinate2D leg_ul = new(7.125, 4);

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

                //Add rectangle around legend
                Coordinate2D legrec_ll = new(7, 1.125);
                Coordinate2D legrec_ur = new(10.875, 4.125);
                ArcGIS.Core.Geometry.Envelope legrecEnv = EnvelopeBuilderEx.CreateEnvelope(legrec_ll, legrec_ur);

                CIMStroke legRecStroke = SymbolFactory.Instance.ConstructStroke(ColorFactory.Instance.BlackRGB);
                CIMPolygonSymbol legrecPolySym = SymbolFactory.Instance.ConstructPolygonSymbol(ColorFactory.Instance.CreateRGBColor(0, 0, 0, 0), SimpleFillStyle.Solid, legRecStroke);
                CIMGraphic rect = GraphicFactory.Instance.CreateSimpleGraphic(legrecEnv, legrecPolySym);

                var legRectInfo = new ElementInfo()
                {
                    Anchor = Anchor.BottomLeftCorner,
                };

                var legendRectangle = ElementFactory.Instance.CreateGraphicElement(newLayout, rect, "Legend Rectangle", true, legRectInfo);


                //Add Map Frame Description automated text to bottom right of layout
                string description = $@"<dyn type =""mapFrame"" name =""Core Map Frame"" property =""description"" />";

                Coordinate2D description_ll = new(7, 0.125);
                Coordinate2D description_ur = new(10.875, 1.125);
                ArcGIS.Core.Geometry.Envelope descEnvPosition = EnvelopeBuilderEx.CreateEnvelope(description_ll, description_ur);

                CIMTextSymbol cimDesc = SymbolFactory.Instance.ConstructTextSymbol(ColorFactory.Instance.BlackRGB, 10, "Arial", "Regular");
                cimDesc.HorizontalAlignment = HorizontalAlignment.Center;
                cimDesc.OffsetY = -5;

                var descInfo = new ElementInfo()
                {
                    Anchor = Anchor.BottomLeftCorner,

                };


                var descText = ElementFactory.Instance.CreateTextGraphicElement(newLayout, TextType.RectangleParagraph, descEnvPosition.Extent, cimDesc, description, "Core Frame Description", true, descInfo);

                if (descText.GetGraphic() is CIMParagraphTextGraphic cimDescText)
                {
                    cimDescText.Frame.BorderSymbol.Symbol.SetSize(1);
                    cimDescText.Frame.BorderSymbol.Symbol.SetColor(ColorFactory.Instance.BlackRGB);


                    descText.SetGraphic(cimDescText);
                }

                //Remove Service Layer
                string serviceLayer = $@"<dyn type =""layout"" name =""Landscape Layout with Inset"" property = ""serviceLayerCredits""/>";
                Coordinate2D serviceLayer_ll = new(0, 0);

                CIMTextSymbol cimServiceLayer = SymbolFactory.Instance.ConstructTextSymbol(ColorFactory.Instance.CreateRGBColor(255, 255, 255, 0), 10);
                var servText = ElementFactory.Instance.CreateTextGraphicElement(newLayout, TextType.PointText, serviceLayer_ll.ToMapPoint(), cimServiceLayer, serviceLayer, "Invisible Service Layer");


                // Return the updated layout
                return newLayout;
            });

            // Open new layout on the GUI thread
            await ProApp.Panes.CreateLayoutPaneAsync(lyt);
        }
    }
}
