using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

/// <summary>
/// Middle layer Object, communicates with data layer and returns types, that are intended to be used
/// by presentation layer.
/// Deals with data associated with Cuisine table
/// </summary>
public class Cuisine
{
    public int CuisineID { get; set; }
    public String CuisineRegion { get; set; }
    public String CuisineName { get; set; }
    private DatabaseConnection dbConn;



    
    /// <summary>
    /// Default constructor - creates new DatabaseConnection object and assigns it to appropriate field. 
    //DataBaseConnection forms basis for communication with data layer.
    /// </summary>
    public Cuisine()
    {
        dbConn = new DatabaseConnection();
    }


    /// <summary>
    /// Method delegates to data layer and returns DataTable with all cuisines
    /// </summary>
    /// <returns>DataTable containing all available cuisines</returns>
    public DataTable getAllCuisiness()
    {
        String command = "SELECT CuisineID, CONCAT(CuisineRegion, ': ', CuisineName) AS CombinedName FROM Cuisine ";
        DataTable table = dbConn.executeReader(command);
        return table;
    }

    /// <summary>
    /// Method delegates to data layer and tries to remove desired Cuisine.
    /// In case query was successful returned value is true, otherwise false
    /// </summary>
    /// <returns>bool - indicates whether query was successful or not</returns>
    public bool removeCuisine()
    {
        dbConn.addParameter("@CuisineID", CuisineID);

        string command = "DELETE FROM Cuisine WHERE CuisineID=@CuisineID";

        return dbConn.executeNonQuery(command) > 0;

    }

    /// <summary>
    /// Method delegates to data layer and tries to remove desired Cuisine from all restaurants.
    /// In case query was successful returned value is true, otherwise false
    /// </summary>
    /// <returns>bool - indicates whether query was successful or not</returns>
    public bool removeCuisineFromRestaurants()
    {
        dbConn.addParameter("@CuisineID", CuisineID);
        string command = "DELETE FROM RestaurantCuisine WHERE CuisineID=@CuisineID";
        return dbConn.executeNonQuery(command) > 0;

    }

    // <summary>
    /// Method delegates to data layer and tries to add new Cuisine to table with available Cuisines.
    /// In case query was successful returned value is true, otherwise false
    /// </summary>
    /// <returns>bool - indicates whether query was successful or not</returns>
    public bool addCuisine()
    {
        dbConn.addParameter("@CuisineName", CuisineName);
        dbConn.addParameter("@CuisineRegion", CuisineRegion);
        String command = "INSERT INTO Cuisine (CuisineName, CuisineRegion) VALUES (@CuisineName, @CuisineRegion)";

        return dbConn.executeNonQuery(command) > 0;
    }



}