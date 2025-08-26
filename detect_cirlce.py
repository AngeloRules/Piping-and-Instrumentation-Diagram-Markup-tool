 
import cv2 
import numpy as np 
  
img = cv2.imread('test_run\Binder-Unit 3400 PEFS 1-24.pdf_20.jpg', cv2.IMREAD_COLOR) 
  

gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY) 
  
gray_blurred = cv2.blur(gray, (7, 7)) 

detected_circles = cv2.HoughCircles(gray_blurred,  
                   cv2.HOUGH_GRADIENT, 1.5, 10, param1 = 60 , 
               param2 = 30, minRadius = 22, maxRadius = 25) 
  
if detected_circles is not None: 
    detected_circles = np.uint16(np.around(detected_circles)) 
    for pt in detected_circles[0, :]: 
        a, b, r = pt[0], pt[1], pt[2] 
        top_left = (a - r, b - r)
        bottom_right = (a + r, b + r)
        cv2.circle(img, (a, b), r, (0, 255, 0), 2) 
        cv2.rectangle(img, top_left, bottom_right, (0, 0, 255), 2)
        cv2.circle(img, (a, b), 1, (0, 0, 255), 3) 
cv2.imshow("Detected Circle", img) 
cv2.imwrite('houghcircle.png', img)
cv2.waitKey(0) 