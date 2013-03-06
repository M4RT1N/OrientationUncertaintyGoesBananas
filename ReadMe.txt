
The VBA application orientation_samplespace_stereogram_with_omega_version_1_0 is
developed from the theories presented in Stigsson M. and Munier R., 2013. 
Orientation Uncertainty goes Bananas: An algorithm to visualise the uncertainty
sample space on stereonets for oriented objects measured in boreholes. Computers
and Geosciences,


RUNING APPLICATION

1) Open the Microsoft Office Excel 2007 application 
   orientation_samplespace_stereogram_with_omega_version_1_0x.xlb
2) Fill the desired values in cells C5 to D9, see definitions section and 
   restrictions section
3) Tick appropriate boxes, see options section
4) Click the 'Calculate and plot' button (Depending on the complexity it might
   take some time to calculate, and progress is shown below the button)


DEFINITIONS

Bearing:     the angle between North and the borehole trajectory projected to
             the horizontal. The angle is measured clockwise from north and has 
             a value between 0° and 360°.

Inclination: the acute angle between the horizontal plane and the trajectory of 
             the borehole. The value of the inclination can be between -90° and 
             90°, where I < 0° corresponds to a borehole pointing downwards.

Alpha angle: the acute dihedral angle between the fracture plane and the 
             trajectory of the borehole. The angle is restricted to be between 
             0° and 90°, where 90° corresponds to a fracture perpendicular to 
             the borehole, i.e. the trajectory of the borehole is parallel to 
             the normal vector of the plane.

Beta angle:  the angle from a reference line (in this application defined as the
             line of the top of the roof of the borehole profile) to the lower 
             inflexion point of the fracture trace on the borehole wall, i.e. 
             where the perimeter of the borehole is the tangent of the fracture 
             trace. The angle is measured clockwise looking in the direction of 
             the borehole trajectory and can hence be between 0° and 360°.

Omega:       the dihedral angle between expected orientation of the fracture and
             the sample space point furthest away from the expected orientation 
             of the fracture.

Grey circle: the expected orientation of the borehole 

Blue circle: the expected orientation of the fracture 



RESTRICTIONS

Valid ranges of values

  0 <= expected alpha          <=  90
  0 <  alpha uncertainty       <=  90
  0 <= expected beta           <= 360
  0 <  beta uncertainty        <= 180
  0 <= expected bearing        <= 360
  0 <= bearing uncertainty     <= 180
-90 <= expected inclination    <=  90
  0 <= inclination uncertainty <=  90

If ticked; 1 <= Nof Dots <= 32000


OPTIONS

Draw Sample Space: Draws the boundary of the uncertainty sample space on the 
             lower hemisphere equal area stereonet as a red line

Draw Omega circle: Draws a blue circle around the expected orientation of the 
             fracture with the dihedral angle Omega

Plot ## random points: Plots the desired number of green points using a 
             rectangular distribution for each of the 4 uncertainty values

Draw the 4 constituents: Draws the sample space for each of the 4 constituents: 
             red solid line: alpha sample space of fracture
             blue solid line: beta sample space of fracture
             yellow dotted line: bearing sample space of borehole
             yellow solid line:  bearing sample space of fracture
             purple dotted line: inclination sample space of borehole
             purple solid line:  inclination sample space of fracture

Draw alpha/beta uncertainties using 9 Bearing/Inclination couples: plots the 
             boundary of the sample space for the alpha and beta values as green
             lines using the 9 combinations from +-uncertainty together with the
             expected value for the bearing and inclination. 
