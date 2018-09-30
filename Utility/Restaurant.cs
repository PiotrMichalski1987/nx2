using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

/// <summary>
/// Middle layer Object, communicates with data layer and returns types, that are intended to be used
/// by presentation layer.
/// Deals with data associated with Restaurant table
/// </summary>
public class Restaurant
{
    public String RestaurantName { set; get; }
    public int RestaurantID { set; get; }

    public int CuisineID { set; get; }

    private DatabaseConnection dbConn;


    /// <summary>
    /// Default constructor - creates new DatabaseConnection object and assigns it to appropriate field. 
    //DataBaseConnection forms basis for communication with data layer.
    /// </summary>
	public Restaurant()
	{
        dbConn = new DatabaseConnection();
	}



    /// <summary>
    /// Method delegates to data layer and tries to remove Cuisine, from restaurant.
    /// In case query was successful returned value is true, otherwise false
    /// </summary>
    /// <returns>bool - indicates whether query was successful or not</returns>
    public bool removeCuisineFromRestaurant()
    {
        dbConn.addParameter("@RestaurantID", RestaurantID);
        dbConn.addParameter("@CuisineID", CuisineID);
        string command = "DELETE FROM RestaurantCuisine WHERE CuisineID=@CuisineID AND RestaurantID=@RestaurantID";
        return dbConn.executeNonQuery(command) > 0;
    }


    /// <summary>
    /// Method delegates to data layer and returns DataTable with all restaurants
    /// </summary>
    /// <returns>DataTable containing all restaurantss, can be null</returns>
    public DataTable getAllRestaurants()
    {
        String command = "Select RestaurantID,RestaurantName FROM Restaurant";
        DataTable table = dbConn.executeReader(command);
        return table;
    }



    // <summary>
    /// Method delegates to data layer and returns DataTable with chefs in selected restaurant
    /// </summary>
    /// <returns>DataTable containing chefs in specific restaurant, can be null</returns>
    public DataTable getChefsInRestaurant()
    {
        dbConn.addParameter("@RestaurantID", RestaurantID);
        dbConn.addParameter("@RoleID", 2);
        String command = "SELECT EmailAddress, RealName FROM RestaurantUser WHERE RestaurantID=@RestaurantID AND RoleID=@RoleID ";

         DataTable table = dbConn.executeReader(command);
        return table;

    }

    /// <summary>
    /// Method delegates to data layer and returns DataTable with all restaurants,except one with specified RestaurantID
    /// </summary>
    /// <returns>DataTable containing all restaurants except one with selected RestaurantID, can be null</returns>
    public DataTable getRemainingRestaurants()
    {
        dbConn.addParameter("@RestaurantID", RestaurantID);

        String command = "Select RestaurantID,RestaurantName FROM Restaurant WHERE RestaurantID <> @RestaurantID ";
        DataTable table = dbConn.executeReader(command);
        return table;
    }


    /// <summary>
    /// Returns name of specific restaurant
    /// 
    /// </summary>
    /// <returns>String, in cases RestaurantID where can not be found, returns empty string</returns>
    public String getRestaurantName()
    {
        dbConn.addParameter("@RestaurantID", RestaurantID);
        String command = "SELECT RestaurantName FROM Restaurant WHERE RestaurantID=@RestaurantID ";
        String result;
        DataTable dt = dbConn.executeReader(command);


        if (dt.Rows.Count < 0)
        {
            result = "";
        }
        else
        {
            result = dt.Rows[0]["RestaurantName"].ToString();
        }

        return result;
    }

    /// <summary>
    /// Returns table with all cuisines in specific restaurant
    /// 
    /// </summary>
    /// <returns>DataTable with cuisines in restaurant with specific RestaurantID, can be null</returns>
    public DataTable selectAllCuisinesInRestaurant()
    {
        dbConn.addParameter("@RestaurantID", RestaurantID);
        String command = "SELECT CuisineID FROM RestaurantCuisine WHERE RestaurantID=@RestaurantID";

        DataTable dt = dbConn.executeReader(command);
        DataTable dt2 = new DataTable();
        foreach(DataRow elem in dt.Rows)
        {
            dbConn.addParameter("@CuisineID", elem["CuisineID"].ToString());          
            DataTable tmp = dbConn.executeReader("SELECT CuisineID, CuisineRegion, CuisineName , CONCAT(CuisineRegion, ': ', CuisineName) AS CombinedName FROM Cuisine WHERE CuisineID=@CuisineID") ;         
            dt2.Merge(tmp);
        }

        return dt2;
    }

    /// <summary>
    /// 
    /// Method delegates to data layer and tries to add new Cuisine to restaurant with specific RestaurantID .
    /// In case query was successful returned value is true, otherwise false
    /// </summary>
    /// <returns>bool - indicates whether query was successful or not</returns>
   
    public bool addCuisineToRestaurant()
    {
        dbConn.addParameter("@RestaurantID", RestaurantID);
        dbConn.addParameter("@CuisineID", CuisineID);
        string command = "INSERT INTO RestaurantCuisine (CuisineID, RestaurantID) VALUES (@CuisineID, @RestaurantID) ";
        return dbConn.executeNonQuery(command) > 0;
    }


    /// <summary>
    /// Method delegates to data layer and returns DataTable with all waiters in specific restaurant
    /// </summary>
    /// <returns>DataTable containing waiters in selected restaurant, can be null</returns>
    public DataTable selectWaitersInRestaurant()
    {
        dbConn.addParameter("@RestaurantID", RestaurantID);
        dbConn.addParameter("@RoleID", 1);

        String command = "Select UserID,UserName FROM RestaurantUser WHERE RestaurantID=@RestaurantID AND RoleID=@RoleID";
        DataTable table = dbConn.executeReader(command);
        return table;
    }


}