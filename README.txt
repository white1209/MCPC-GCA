The MCPC Gospel Car Arrangement System is an automated transport-planning application designed to efficiently organize Sunday service transportation for church members. Built using Python and Streamlit, the system integrates Google Sheets, OpenRouteService (ORS), and intelligent routing logic to generate a clear, structured transport arrangement with minimal manual effort.

The system retrieves participant data directly from a Google Sheet, normalizes pickup locations, and maps them to predefined geographic coordinates. Using OpenRouteService’s distance matrix, it calculates optimal driving routes and determines the most efficient pickup sequence that minimizes total travel distance.

To ensure smooth operations, the system intelligently:
1. Separates worship enablers into an early trip with adjusted departure times
2. Groups passengers by pickup location and vehicle capacity
3. Automatically removes conflicts 
4. Prioritizes specific locations when arranging routes
5. Estimates travel times and generates clear estimated time arrived
6. Handles departure trips, carpool arrangements, and after-service transportation

Finally, the system produces a well-formatted transport brief, including trip sequences, pickup times, passenger lists, and return schedules. The output is displayed directly within the Streamlit interface, making it easy for coordinators to review, copy, and share.

Overall, this system significantly reduces coordination effort, minimizes human error, and ensures a smooth, organized transportation experience for church services and fellowship activities.
