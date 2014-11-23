using System;
using System.Collections.Generic;
using System.IO;


using System.Windows.Forms;
using System.Drawing;

using System.Windows.Media.Imaging;
using System.Reflection;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Structure;

namespace OpenDraftingToolkit {


    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]
    public class LeaderAdjustment : IExternalCommand, IExternalApplication {

        private const double LEADER_LENGTH_THRESHOLD = 0.125 / 12;
        private const double DELTA_Y_THRESHOLD = 0.125 / 12;

        private const string PRESET_30 = "30_degrees";
        private const string PRESET_45 = "45_degrees";
        private const string PRESET_60 = "60_degrees";


        public static double _leader_angle;
        public static double _combo_box_angle;

        static string TRANSACTION_NAME = "Leader Adjustment";

        //leader elbow in text box
        public static FailureDefinitionId _warn_elbow_in_text_id;
        public static FailureDefinitionId _warn_elbow_too_short_id;

        //leader too short warnings
        private FailureDefinition _warn_elbow_in_text_obj;
        private FailureDefinition _warn_elbow_too_short_obj;


        #region "IExternalApplication Defintions"

        /// <summary>
        /// Application Startup - Define Failures and setup UI
        /// </summary>
        /// <param name="app"></param>
        /// <returns></returns>
        public Result OnStartup(UIControlledApplication app) {

            _combo_box_angle = 60;

            try {

                /*
                 *   Failure Defintions
                 * 
                 */

                // Create failure definition Ids
                Guid w1_guid = new Guid("83906ecb-f179-4d24-9477-b4f25834bbd3");
                _warn_elbow_in_text_id = new FailureDefinitionId(w1_guid);

                // Create failure definitions and add resolutions
                _warn_elbow_in_text_obj = FailureDefinition.CreateFailureDefinition(_warn_elbow_in_text_id, FailureSeverity.Warning, "Generated elbow point lands in the text: move the text to a new position.");

                _warn_elbow_in_text_obj.AddResolutionType(FailureResolutionType.SkipElements, "SkipElements", typeof(DeleteElements));

                Guid w2_guid = new Guid("b1adddfc-0221-4d03-b9c4-72dc875735c0");
                _warn_elbow_too_short_id = new FailureDefinitionId(w2_guid);

                _warn_elbow_too_short_obj = FailureDefinition.CreateFailureDefinition(_warn_elbow_too_short_id, FailureSeverity.Warning, "Generated arrow segement is too short: move elobw or text to a new position");

                /*
                 *   Build UI Elements
                 * 
                 */

                //location of executing assembly
                string assembly_path = Assembly.GetExecutingAssembly().Location;
                string assembly_directory = Path.GetDirectoryName(assembly_path);

                //define ribbon panel
                RibbonPanel ribbonPanel = app.CreateRibbonPanel("Leader Adjustment");

                PushButtonData buttonData = new PushButtonData("cmdLeaderAdjustment", "Adjust Leaders", assembly_path, "OpenDraftingToolkit.LeaderAdjustment");

                PushButton pushButton = ribbonPanel.AddItem(buttonData) as PushButton;


                string lrg_img_path = assembly_directory + Path.DirectorySeparatorChar + "LeaderAdjustmentIcon.png";

                Uri bmuri = new Uri(lrg_img_path);
                BitmapImage bm = new BitmapImage(bmuri);

                pushButton.LargeImage = bm;
                pushButton.Image = bm;
                pushButton.ToolTip = "Sets the elbow angle of Text leaders to a fixed angle";

                Autodesk.Revit.UI.ComboBoxData cbd = new ComboBoxData("Angle Presets");

                cbd.ToolTip = "Sets the Leader Angle";
                cbd.LongDescription = "This is the Angle the Leader elbow will be set to.";

                Autodesk.Revit.UI.ComboBox cb = ribbonPanel.AddItem(cbd) as Autodesk.Revit.UI.ComboBox;
                cb.ItemText = "Leader Angle Presets";


                cb.AddItem(new ComboBoxMemberData(PRESET_60, "Leader Angle: 60 degrees"));
                cb.AddItem(new ComboBoxMemberData(PRESET_45, "Leader Angle: 45 degrees"));
                cb.AddItem(new ComboBoxMemberData(PRESET_30, "Leader Angle: 30 degrees"));

                //add the handler
                cb.CurrentChanged += new EventHandler<Autodesk.Revit.UI.Events.ComboBoxCurrentChangedEventArgs>(cb_CurrentChanged);

                //set the intial angle manually
                _combo_box_angle = 60;

                _leader_angle = Math.PI * 0.25; //45 degrees


                return Result.Succeeded;

            } catch (Exception ex) {

                //System.Diagnostics.Debug.Print("*** plugin OnStartup Failed! :: " + ex.Message + " ***");
                return Result.Failed;

            }



        }

        /// <summary>
        /// Event Handler for ComboBox changes
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void cb_CurrentChanged(object sender, Autodesk.Revit.UI.Events.ComboBoxCurrentChangedEventArgs e) {

            try {

                Autodesk.Revit.UI.ComboBox cb = (Autodesk.Revit.UI.ComboBox)sender;
                ComboBoxMember cm = cb.Current;

                string cm_id = cm.Name;

                if (cm_id == PRESET_60) {
                    _combo_box_angle = 60;
                } if (cm_id == PRESET_45) {
                    _combo_box_angle = 45;
                } if (cm_id == PRESET_30) {
                    _combo_box_angle = 30;
                }

                string valstr = cm.ItemText;


            } catch (Exception ex) {

                System.Diagnostics.Debug.Print("Error in ComboBox Event: " + ex.Message);


            }



        }//end OnStartup


        /// <summary>
        /// Sutdown Event - nothing to do here.
        /// </summary>
        /// <param name="app"></param>
        /// <returns></returns>
        public Result OnShutdown(UIControlledApplication app) {

            return Result.Succeeded;

        }//end OnShutdown


        #endregion

        #region "IExternalCommand"


        /// <summary>
        /// Starts the interactive Command to Adjust Leaders
        /// </summary>
        /// <param name="commandData"></param>
        /// <param name="message"></param>
        /// <param name="elements"></param>
        /// <returns></returns>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements) {
            //Get application and document objects
            UIApplication uiApp = commandData.Application;
            Document doc = uiApp.ActiveUIDocument.Document;

            Autodesk.Revit.DB.View view = uiApp.ActiveUIDocument.ActiveView;


            double angle_degrees = 0;
            double angle_radians = 0;

            //so the scale in Revit is stored as an Integer, who knew?....
            double view_scale = view.Scale;


            //Pick a text Entity
            Selection sel = uiApp.ActiveUIDocument.Selection;
            Reference elemRef = null;
            Element element = null;

            bool do_get = true;

            while (do_get) {

                try {


                    elemRef = sel.PickObject(ObjectType.Element, "Select a Text for Leader adjustment");
                    element = doc.GetElement(elemRef);

                    //do this here to pick up changes by user
                    angle_degrees = _combo_box_angle;
                    angle_radians = (Math.PI / 180.0) * angle_degrees;


                } catch (Autodesk.Revit.Exceptions.OperationCanceledException ex) {

                    //make sure this doesn't invoke autorollback
                    //want to leave previous changes intact
                    return Result.Succeeded;

                } catch (Exception ex) {

                    return Result.Failed;

                }

                //don't nest this try/catch block so exeptions from this section can be isolated easily.

                try {

                    TextNote textObj = null;
                    //check element
                    if (element == null) throw new Exception("Invalid null Element in Leader Adjustment");

                    //init transaction
                    Transaction trans = new Transaction(doc);
                    trans.Start(TRANSACTION_NAME);

                    //try cast
                    try {

                        textObj = element as TextNote;

                    } catch {

                        //not a text element
                        //let it fall through.

                    }

                    if (textObj != null) {

                        //get the textType
                        TextElementType tfs = textObj.Symbol;

                        /*
                         *   Define Text Bounding Box (not Revit BoundingBox wich includes Leaders)
                         * 
                         */
                        Parameter pad_param = tfs.get_Parameter(BuiltInParameter.LEADER_OFFSET_SHEET);
                        double pad = pad_param.AsDouble();
                        pad *= view_scale;

                        XYZ textCoord = textObj.Coord;

                        double width = textObj.Width;
                        double height = textObj.Height;
                        double posx = textCoord.X;
                        double posy = textCoord.Y;
                        double xmin = 0;
                        double xmax = 0;
                        double ymin = posy - (height * view_scale);
                        double ymax = ymin + (height * view_scale);

                        ymin -= pad;
                        ymax += pad;

                        Parameter halign = textObj.get_Parameter(BuiltInParameter.TEXT_ALIGN_HORZ); // ("Horizontal Align");

                        //not sure best way to do this...
                        TextAlignFlags align_code = (TextAlignFlags)halign.AsInteger();

                        switch (align_code) {

                            case TextAlignFlags.TEF_ALIGN_LEFT:

                                xmin = posx - pad;
                                xmax = posx + width - pad;
                                break;


                            case TextAlignFlags.TEF_ALIGN_CENTER:

                                xmin = posx - (width * 0.5);
                                xmax = posx + (width * 0.5);
                                break;

                            case TextAlignFlags.TEF_ALIGN_RIGHT:

                                xmin = posx - width + pad;
                                xmax = posx;
                                break;

                            default:

                                throw new Exception("Invalid Text Horizontal Align Flag in Leader Adjustment: " + halign);

                        }

                        /*
                         *   Process Leaders
                         * 
                         */

                        LeaderArray leaders = textObj.Leaders;
                        int lc = leaders.Size;


                        for (int l = 0; l < lc; l++) {

                            Leader leader = leaders.get_Item(l);

                            //copy points...
                            XYZ elbow = new XYZ(leader.Elbow.X, leader.Elbow.Y, 0);
                            XYZ target = new XYZ(leader.End.X, leader.End.Y, 0);


                            if (calcNewElbow(angle_radians, view_scale, ref target, ref elbow)) {

                                //check x bounds... elbow cannot land in xbounds
                                double ex = elbow.X;
                                double ey = elbow.Y;

                                //make sure elbow point doesn't fall in text box
                                //Revit actually seems to handle this well, but doesn't give intended result: you get straight or displaced leaders
                                if (!(ex >= xmin && ex <= xmax && ey >= ymin && ey <= ymax)) {

                                    //everything is ok so update both points in case leader is straightened
                                    leader.Elbow = elbow;
                                    leader.End = target;


                                } else {

                                    FailureMessage fm = new FailureMessage(_warn_elbow_in_text_id);
                                    fm.SetFailingElement(element.Id);
                                    doc.PostFailure(fm);

                                }

                            } else {


                                FailureMessage fm = new FailureMessage(_warn_elbow_too_short_id);
                                fm.SetFailingElement(element.Id);
                                doc.PostFailure(fm);

                            }

                        }//end foreach


                        //commit trans after leaders are adjusted
                        trans.Commit();

                    }//end if textObj null

                    trans.Dispose();
                    trans = null;

                } catch (Exception ex) {

                    MessageBox.Show("An unexpected error occured while Adjusting Leaders: " + ex.Message, "Error Adjusting Leaders", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return Result.Failed;
                }



            }//end while


            return Result.Succeeded;

        }//end Execute


        #endregion


        #region "Calculations"
        /// <summary>
        /// Calculates a new ebow position relative to the leader end (arrow)
        /// </summary>
        /// <param name="angle">Angle of arrow segment - this assumes elbow segment is horizontal</param>
        /// <param name="view_scale">Scale of the active view</param>
        /// <param name="x_limit"></param>
        /// <param name="target"></param>
        /// <param name="elbow"></param>
        /// <returns></returns>
        private bool calcNewElbow(double angle, double view_scale, ref XYZ target, ref XYZ elbow) {

            //arrow.
            bool success = true;

            double delta_x = 0;
            double delta_y = 0;

            double pi = (float)Math.PI;
            double hpi = Math.PI * 0.5;
            double theta = 0;
            double target_angle = 0;

            //deltas
            delta_x = elbow.X - target.X;
            delta_y = elbow.Y - target.Y;

            //if there is insufficient change in y - straighten leader...
            if (Math.Abs(delta_y) < DELTA_Y_THRESHOLD * view_scale) {

                XYZ new_target = new XYZ(target.X, elbow.Y, 0);

                double tdist = new_target.DistanceTo(elbow);

                if (tdist < LEADER_LENGTH_THRESHOLD * view_scale) {

                    return false;

                } else {
                    //adjust the target 
                    target = new_target;

                    return true;
                }
            }




            theta = Math.Atan2(delta_y, delta_x);

            //find quadrant and adjust angle
            if (theta >= 0 && theta < hpi) {

                target_angle = angle;
                //System.Diagnostics.Debug.Print("quadrant 1");

            } else if (theta >= hpi && theta < pi) {

                target_angle = pi - angle;
                // System.Diagnostics.Debug.Print("quadrant 2");

            } else if (theta >= -pi && theta < -hpi) {

                target_angle = pi + angle;
                //System.Diagnostics.Debug.Print("quadrant 3");

            } else if (theta < 0 && theta >= -hpi) {

                target_angle = (2 * pi) - angle;
                // System.Diagnostics.Debug.Print("quadrant 4");
            }

            double abs_dy = Math.Abs(delta_y);
            double len = abs_dy / Math.Sin(angle);
            // float hyp = len;//sqrt (pow(abs(delta_x),2) + pow(abs(delta_y),2));



            double nx = len * Math.Cos(target_angle);
            double ny = len * Math.Sin(target_angle);

            XYZ new_elbow = new XYZ(target.X + nx, target.Y + ny, 0);

            double dist = new_elbow.DistanceTo(target);
            //make sure it's not too short
            if (dist < LEADER_LENGTH_THRESHOLD * view_scale) {

                //MessageBox.Show("Generated arrow segment is too short - move elbow closer to text if possible.", "Leader Adjustment Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;

            }


            //set the new elbow
            elbow = new_elbow;


            return success;

        }

        #endregion

    }//end class

}//end namespace