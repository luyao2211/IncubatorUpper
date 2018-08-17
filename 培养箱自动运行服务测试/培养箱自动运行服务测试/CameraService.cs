using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using CSharp_OPTControllerAPI;
using Basler.Pylon;

namespace CameraService
{
    class CameraServiceAPI
    {
        public void Shoot_One_Slot(string filename, string ftp_addr, bool is_up, string PhotoDir)
        {
            ftp_addr = "";
            if (!Directory.Exists(PhotoDir))
            {
                Directory.CreateDirectory(PhotoDir);
            }

            String ip_OPT = "192.168.5.162";

            int intensity = 5;
            int channel = 1; //channel 1: up channel2:down

            String Camera_SN_lower = "22604711";
            String Camera_SN_upper = "22454039";

            //String parafile = "D:\\20180516.txt";

            try
            {
                if(is_up)
                {
                    Turn_On_LED(ip_OPT, 1, intensity);
                    Grab_One_Image_v3(Camera_SN_upper, filename);
                    //Call localhostserver for analysis and upload results
                    Turn_Off_LED(ip_OPT, 1);
                }
                else
                {
                    Turn_On_LED(ip_OPT, 2, intensity);
                    Grab_One_Image_v3(Camera_SN_lower, filename); //TODO: Camera_SN_up as substitution temporarily
                    //Call localhostserver for analysis and upload results
                    Turn_Off_LED(ip_OPT, 2);
                }

            }
            catch (Exception ee)
            {
                //exitCode = 1;
                Console.Error.WriteLine("\nProcess Encounters Error!");
            }
            finally
            {
                //Console.Error.WriteLine("\nPress enter to exit.");
                //Console.ReadLine();
            }
        }
        public static void Turn_On_LED(String IPAddr, Int32 channel, int intensity)
        {
            long lRet = -1;
            OPTControllerAPI OPTController = new OPTControllerAPI();
            if (IPAddr == "")
            {
                Console.WriteLine("\nIP Address is not regular!");
                return;
            }
            lRet = OPTController.CreateEtheConnectionByIP(IPAddr);
            if (lRet != 0)
            {
                Console.WriteLine("\nFail to connect by IP");
                return;
            }
            else
            {
                if (OPTController.TurnOnChannel(channel) == 0)
                {
                    Console.WriteLine("\nChannel Turned On successfully!");
                    if (OPTController.SetIntensity(channel, intensity) == 0)
                    {
                        Console.WriteLine("Set intensity successfully");
                    }
                    else
                    {
                        Console.WriteLine("Fail to set intensity");
                        return;
                    }
                }
                else
                {
                    Console.WriteLine("\nChannel Failed to Turned on!");
                    return;
                }
                lRet = OPTController.DestoryEtheConnect();
                if (0 != lRet)
                {
                    Console.WriteLine("Failed to disconnect Ethernet connection by IP");
                    return;
                }
                else
                {
                    Console.WriteLine("Successfully disconnected Ethernet connection by IP");
                }
            }
            return;
        }
        public static void Turn_Off_LED(String IPAddr, Int32 channel)
        {
            long lRet = -1;
            OPTControllerAPI OPTController = new OPTControllerAPI();
            if (IPAddr == "")
            {
                Console.WriteLine("\nIP Address is not regular!");
                return;
            }
            lRet = OPTController.CreateEtheConnectionByIP(IPAddr);
            if (lRet != 0)
            {
                Console.WriteLine("\nFail to connect by IP");
                return;
            }
            else
            {
                if (OPTController.TurnOffChannel(channel) == 0)
                {
                    Console.WriteLine("\nChannel Turned Off successfully!");
                    lRet = OPTController.DestoryEtheConnect();
                    if (0 != lRet)
                    {
                        Console.WriteLine("Failed to disconnect Ethernet connection by IP");
                        return;
                    }
                    else
                    {
                        Console.WriteLine("Successfully disconnected Ethernet connection by IP");
                    }
                }
                else
                {
                    Console.WriteLine("\nChannel Failed to Turned off!");
                    return;
                }

            }
        }

        public static void Grab_One_Image_v3(String divece_SN, String filename)
        {
            bool camera_on_list = false;
            Camera camera = null;
            ICameraInfo tag = null;
            try
            {
                // Ask the camera finder for a list of camera devices.
                List<ICameraInfo> allCameras = CameraFinder.Enumerate();


                // Check if the given camera is in the camera list
                foreach (ICameraInfo cameraInfo in allCameras)
                {
                    if (cameraInfo[CameraInfoKey.SerialNumber] == divece_SN)
                    {
                        //save tags for camera
                        camera_on_list = true;
                        tag = cameraInfo;
                        break;
                    }
                }

                if (!camera_on_list)
                {
                    Console.WriteLine("No such camera: " + divece_SN + "\n");
                    return;
                }

                if (camera_on_list == true && tag != null)
                {
                    try
                    {
                        //create a new camera object
                        using (camera = new Camera(tag))
                        {
                            camera.CameraOpened += Configuration.AcquireContinuous;

                            // Register for the events of the image provider needed for proper operation.
                            //camera.ConnectionLost += OnConnectionLost;
                            //camera.CameraOpened += OnCameraOpened;
                            //camera.CameraClosed += OnCameraClosed;
                            //camera.StreamGrabber.GrabStarted += OnGrabStarted;
                            //camera.StreamGrabber.ImageGrabbed += OnImageGrabbed;
                            //camera.StreamGrabber.GrabStopped += OnGrabStopped;

                            // Open the connection to the camera device
                            camera.Open();

                            //Set the parameters for the controls.
                            camera.Parameters[PLCamera.GainAuto].SetValue("Off");
                            camera.Parameters[PLCamera.GainRaw].SetValue(85);
                            camera.Parameters[PLCamera.DecimationHorizontal].SetValue(2);
                            camera.Parameters[PLCamera.DecimationVertical].SetValue(2);
                            camera.Parameters[PLCamera.BalanceWhiteAuto].SetValue("Off");
                            camera.Parameters[PLCamera.LightSourceSelector].SetValue("Daylight");
                            camera.Parameters[PLCamera.BalanceRatioSelector].SetValue("Red");
                            camera.Parameters[PLCamera.BalanceRatioRaw].SetValue(101);
                            camera.Parameters[PLCamera.PixelFormat].SetValue("YUV422_YUYV_Packed");
                            camera.Parameters[PLCamera.ExposureTimeRaw].SetValue(24500);

                            // Start Stream Grabber
                            try
                            {
                                // Starts the grabbing of one image.
                                camera.Parameters[PLCamera.AcquisitionMode].SetValue(PLCamera.AcquisitionMode.SingleFrame);
                                camera.StreamGrabber.Start();
                                IGrabResult grabResult = camera.StreamGrabber.RetrieveResult(10000, TimeoutHandling.ThrowException);
                                using (grabResult)
                                {
                                    if (grabResult.GrabSucceeded)
                                    {
                                        byte[] buffer = grabResult.PixelData as byte[];
                                        ImagePersistence.Save(ImageFileFormat.Jpeg, filename, grabResult);
                                        Console.WriteLine("Saving... " + filename + "\n");

                                    }
                                    else
                                    {
                                        //error
                                    }
                                }
                                camera.StreamGrabber.Stop();
                                camera.Close();

                            }
                            catch (Exception exception)
                            {
                                ShowException(exception);
                                camera.StreamGrabber.Stop();
                                camera.Close();
                            }


                            //camera.Dispose(); //May damage the buffer, do not use it!
                        }


                    }

                    catch (Exception exception)
                    {
                        ShowException(exception);
                        camera.StreamGrabber.Stop();
                        camera.Close();
                    }
                }
            }

            catch (Exception exception)
            {
                ShowException(exception);
                camera.StreamGrabber.Stop();
                camera.Close();
            }

        }


        private static void ShowException(Exception exception)
        {
            Console.WriteLine("Exception caught:\n" + exception.Message, "Error");
        }
    }
}
