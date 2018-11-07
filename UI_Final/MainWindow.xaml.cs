using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using Microsoft.Kinect;

using System.Runtime.InteropServices;
using System.Windows.Interop;
using System.Timers;

namespace UI_Final
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private KinectSensor newSensor;
        const int skeletonCount = 6;
        private Skeleton[] allSkeletons = new Skeleton[skeletonCount];

        /// <summary>
        /// The following are hex value mouse commands.
        /// </summary>
        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;

        /// <summary>
        /// The offsets are calculated based on the width of the kinect image data versus the desktop.
        /// These may need to change based on the size of the desktop.
        /// </summary>
        private const int XOFFSET = 3;
        private const int YOFFSET = 2;
        private const int CURSOR_OFFSET = 30;

        int potentialClickLocationX = 0;
        int potentialClickLocationY = 0;

        int lastXLocation = 0;
        int lastYLocation = 0;

        bool leftClickOn = false;

        Timer leftClickTimer = new Timer();
        Timer doubleClickTimer = new Timer();

        public MainWindow()
        {
            ///Timer for the left click event.
            leftClickTimer.Interval = 1500;
            leftClickTimer.Elapsed += new ElapsedEventHandler(leftClickTimer_Elapsed);

            ///Timer for the double click event.
            doubleClickTimer.Interval = 3500;
            doubleClickTimer.Elapsed += new ElapsedEventHandler(doubleClickTimer_Elapsed);

            InitializeComponent();
        }

        /// <summary>
        /// Handles the Loaded event of the Window control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs" /> instance containing the event data.</param>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ///Look for an available sensor.
            foreach (var potentialSensor in KinectSensor.KinectSensors)
            {
                if (potentialSensor.Status == KinectStatus.Connected)
                {
                    ///We found a sensor to use!
                    this.newSensor = potentialSensor;
                    break;
                }
            }
            ///If we did not find a Kinect, just exit.
            if (newSensor == null)
            {
                MessageBoxResult result = MessageBox.Show("An error occured.\nCrap, we could not find any Kinect sensors!\nSorry but we need to close down.", "Oops!", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                this.Close();
            }
            ///Initilize the sensors you want to use with the plugged in Kinect.
            ///Experimenting with different resolution settings. You could leave these blank.
            newSensor.DepthStream.Enable(DepthImageFormat.Resolution640x480Fps30);
            newSensor.SkeletonStream.Enable();
            ///Bind the event to occur when the kinect indicates it has frames ready to process.
            newSensor.AllFramesReady += new EventHandler<AllFramesReadyEventArgs>(newSensor_AllFramesReady);
            try
            {
                ///Try and start the kinect
                newSensor.Start();
            }
            catch (System.IO.IOException exc)
            {
                ///An error occured. Inform the user.
                MessageBoxResult result = MessageBox.Show("An error occured.\n" + exc.InnerException.ToString(), "Oops!", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);
                return;
            }
            ///This is the default cursor location
            SetCursorLocation(0, 0);
        }

        /// <summary>
        /// Sets the cursor location relative to the application window.
        /// </summary>
        /// <param name="x">The new x position of the cursor.</param>
        /// <param name="y">The new y position of the cursor.</param>
        void SetCursorLocation(int x, int y)
        {
            var desktopHandle = Win32.GetDesktopWindow();
            Win32.POINT p = new Win32.POINT();
            p.x = x;
            p.y = y;
            ///This is a handle to the current window
            Win32.ClientToScreen((IntPtr)desktopHandle, ref p);
            Win32.SetCursorPos(p.x, p.y);
        }

        /// <summary>
        /// Handles the AllFramesReady event of the newSensor control. (The Kinect)
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="AllFramesReadyEventArgs" /> instance containing the event data.</param>
        void newSensor_AllFramesReady(object sender, AllFramesReadyEventArgs e)
        {
            ///Get the first available skeleton
            Skeleton first = GetFirstSkeleton(e);
            if (first == null)
            {
                return;
            }
            ///Get the camera point for this skeleton.
            GetCameraPoint(first, e);
        }

        /// <summary>
        /// Handles the Elapsed event of the click Timer control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="ElapsedEventArgs" /> instance containing the event data.</param>
        void leftClickTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            ///Check that the location of the mouse is still within a valid location for a click event.
            if (leftClickOn && ((Math.Abs(potentialClickLocationX - lastXLocation) < CURSOR_OFFSET) && Math.Abs(potentialClickLocationY - lastYLocation) < CURSOR_OFFSET))
            {
                ///Click
                Win32.mouse_event(MOUSEEVENTF_LEFTUP, (uint)potentialClickLocationX, (uint)potentialClickLocationY, 0, 0);
                leftClickOn = false;
            }
            else if (!leftClickOn && ((Math.Abs(potentialClickLocationX - lastXLocation) < CURSOR_OFFSET) && Math.Abs(potentialClickLocationY - lastYLocation) < CURSOR_OFFSET))
            {
                ///Click
                Win32.mouse_event(MOUSEEVENTF_LEFTDOWN, (uint)potentialClickLocationX, (uint)potentialClickLocationY, 0, 0);
                leftClickOn = true;
            }
            leftClickTimer.Stop();
        }

        /// <summary>
        /// Handles the Elapsed event of the enableDoubleClick control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="ElapsedEventArgs" /> instance containing the event data.</param>
        void doubleClickTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            ///Check if the mouse location is still within a valid click event range
            if ((Math.Abs(potentialClickLocationX - lastXLocation) < CURSOR_OFFSET) && (Math.Abs(potentialClickLocationY - lastYLocation) < CURSOR_OFFSET))
            {
                Win32.mouse_event(MOUSEEVENTF_LEFTUP, (uint)potentialClickLocationX, (uint)potentialClickLocationY, 0, 0);
                Win32.mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, (uint)potentialClickLocationX, (uint)potentialClickLocationY, 0, 0);
                Win32.mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, (uint)potentialClickLocationX, (uint)potentialClickLocationY, 0, 0);
            }
            doubleClickTimer.Stop();
        }

        /// <summary>
        /// Gets the camera point for a given skeleton.
        /// </summary>
        /// <param name="first">A skeleton that we want the points for.</param>
        /// <param name="e">The <see cref="AllFramesReadyEventArgs" /> instance containing the event data.</param>
        private void GetCameraPoint(Skeleton first, AllFramesReadyEventArgs e)
        {
            using (DepthImageFrame depth = e.OpenDepthImageFrame())
            {
                if (depth == null || newSensor == null)
                {
                    return;
                }
                ///Get the point for the left hand
                DepthImagePoint cursorLeftHandDepthPoint = depth.MapFromSkeletonPoint(first.Joints[JointType.HandLeft].Position);
                ///Get the point for the right hand
                DepthImagePoint cursorRightHandDepthPoint = depth.MapFromSkeletonPoint(first.Joints[JointType.HandRight].Position);
                ///Based on the location of the depth point, set the cursor to that location on the screen.
                CursorLogic(cursorLeftHandDepthPoint, cursorRightHandDepthPoint);
            }
        }

        /// <summary>
        /// Handles hand to cursor translation logic and click timing coordination.
        /// </summary>
        /// <param name="isLeftHand">if set to <c>true</c> [is left hand].</param>
        /// <param name="point">The point.</param>
        public void CursorLogic(DepthImagePoint leftHand, DepthImagePoint rightHand)
        {
            ///Set the cursor location
            SetCursorLocation(rightHand.X * XOFFSET, rightHand.Y * YOFFSET);

            lastXLocation = rightHand.X * XOFFSET;
            lastYLocation = rightHand.Y * YOFFSET;
            ///Check if the offset of the cursor is less then 50 since last time.
            if ((Math.Abs(rightHand.X * XOFFSET - lastXLocation) < CURSOR_OFFSET) && (Math.Abs(rightHand.Y * YOFFSET - lastYLocation) < CURSOR_OFFSET))
            {
                if (!leftClickTimer.Enabled && !doubleClickTimer.Enabled)
                {
                    potentialClickLocationX = rightHand.X * XOFFSET;
                    potentialClickLocationY = rightHand.Y * YOFFSET;
                    leftClickTimer.Start();
                    doubleClickTimer.Start();
                }
            }
            else
            {
                ///Reset the timers
                leftClickTimer.Stop();
                doubleClickTimer.Stop();
            }
        }

        /// <summary>
        /// Stops the kinect.
        /// </summary>
        /// <param name="sensor">The sensor.</param>
        void StopKinect(KinectSensor sensor)
        {
            if (sensor != null)
            {
                sensor.Stop();
                sensor.AudioSource.Stop();
            }
        }

        /// <summary>
        /// Gets the first skeleton available.
        /// </summary>
        /// <param name="e">The <see cref="AllFramesReadyEventArgs" /> instance containing the event data.</param>
        /// <returns></returns>
        private Skeleton GetFirstSkeleton(AllFramesReadyEventArgs e)
        {
            ///The kinect will give us each stream of data as a seperate entity.
            ///This allows us to do things like just get the data for the skeletons.
            ///We could also do this for depth and colour.
            using (SkeletonFrame skeletonFrameData = e.OpenSkeletonFrame())
            {
                if (skeletonFrameData == null)
                {
                    return null;
                }
                ///Add out skeleton to the all skeletons array.
                skeletonFrameData.CopySkeletonDataTo(allSkeletons);
                ///Some qick searching to find the skeleton currently being tracked.
                Skeleton firstSkele = (from skeles in allSkeletons
                                       where skeles.TrackingState == SkeletonTrackingState.Tracked
                                       select skeles).FirstOrDefault();
                ///Return the correct and currently tracked skeleton.
                return firstSkele;
            }
        }

        /// <summary>
        /// Handles the Closing event of the Window control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="System.ComponentModel.CancelEventArgs" /> instance containing the event data.</param>
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            StopKinect(newSensor);
        }
    }

    /// <summary>
    /// This class is basically a bridge that allows us to use the old
    /// windows forms way of doing things to control the mouse cursor.
    /// WPF is not as reliant on User32 so this is why we need this kind of work around here.
    /// See http://stackoverflow.com/questions/6549371/what-is-user32-dll-and-how-it-is-used-in-wpf for more details about User32.
    /// </summary>
    public class Win32
    {
        [DllImport("User32.Dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);

        [DllImport("User32.Dll")]
        public static extern long GetDesktopWindow();

        [DllImport("User32.Dll")]
        public static extern long SetCursorPos(int x, int y);

        [DllImport("User32.Dll")]
        public static extern bool ClientToScreen(IntPtr hWnd, ref POINT point);

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int x;
            public int y;
        }
    }
}
