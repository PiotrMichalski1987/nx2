using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;

//// <summary>
/// Middle layer Object, communicates with data layer and returns types, that are intended to be used
/// by presentation layer.
/// Deals with data associated with RestaurantUser table
/// </summary>
public class User
{
    public String UserName { get; set; }
    public String UserPassword { get; set;}
    public String UserEmail { get; set; }
    public String RealName { get; set; }
    public int RestaurantID { get; set; }
    public int RoleID { get; set; }

    public String Salt { get; set; }


    private DatabaseConnection dbConn;


    // <summary>
    /// Default constructor - creates new DatabaseConnection object and assigns it to appropriate field. 
    //DataBaseConnection forms basis for communication with data layer.
    /// </summary>
    public User()
    {
        dbConn = new DatabaseConnection();
    }



    /// <summary>
    /// Method delegates to data layer and tries to determine whether userName exists or not.
    /// In case query was successful returned value is true, otherwise false
    /// </summary>
    /// <returns>bool - indicates whether query was successful or not</returns>
    public bool userNameExists()
    {
        dbConn.addParameter("@UserName",UserName);

        String command = "SELECT COUNT(UserName)FROM RestaurantUser WHERE UserName=@UserName";

        int result = dbConn.executeScalar(command);

        return result > 0 || result == -1;
    }


    /// <summary>
    /// Method delegates to data layer and authenticates user if given Passord and UserName are correct
    /// </summary>
    /// <returns>returns int: -1 Database error, query could not be performed, 0 not authenticated but not null, 1 user authenticated</returns>
    public int authenticateUser()
    {
        dbConn.addParameter("@UserName", UserName);
        dbConn.addParameter("@UserPassword", UserPassword);


        string command = "SELECT UserID, UserName, RoleID, RestaurantID, RealName, EmailAddress FROM RestaurantUser " +
                        "WHERE UserName=@UserName AND UserPassword=@UserPassword";

        DataTable table = dbConn.executeReader(command);

        if (table == null)
            return -1;

        if (table.Rows.Count > 0)
        {
            HttpContext.Current.Session["UserID"] = table.Rows[0]["UserID"].ToString();
            HttpContext.Current.Session["UserName"] = table.Rows[0]["UserName"].ToString();
           HttpContext.Current.Session["RoleID"] = table.Rows[0]["RoleID"].ToString();
         HttpContext.Current.Session["RestaurantID"] = table.Rows[0]["RestaurantID"].ToString();
            HttpContext.Current.Session["UserEmail"] = table.Rows[0]["EmailAddress"].ToString();

            return 1;
        }
        else
        {
            return 0;
        }
 
    }

    /// <summary>
    /// Method delegates to data layer and returns user Salt if UserID is correct
    ///
    /// </summary>
    /// <returns>String, with salt value, "null" if query failed or "" if table is empty</returns>

    public String getSalt()
    {
        dbConn.addParameter("@UserName", UserName);
        


        string command = "SELECT Salt FROM RestaurantUser WHERE UserName=@UserName";

        DataTable table = dbConn.executeReader(command);
        String result = "";

        if (table == null )
        {
            return "null";
        }
        
       

        if (table.Rows.Count > 0)
        {
            result =table.Rows[0]["Salt"].ToString();

     
        }
        return result;
        

    }


    public bool addUser()
    {
        

        dbConn.addParameter("@UserName", UserName);
        dbConn.addParameter("@UserPassword", UserPassword);
        dbConn.addParameter("@UserEmail", UserEmail);
        dbConn.addParameter("@UserRealName", RealName);     
        dbConn.addParameter("@RestaurantID", RestaurantID);
        dbConn.addParameter("@RoleID", RoleID);
        dbConn.addParameter("@Salt", Salt);

        String command = "INSERT INTO RestaurantUser (UserName, UserPassword,  EmailAddress, RealName, RestaurantID, RoleID, Salt) " +
            "VALUES (@UserName, @UserPassword, @UserEmail, @UserRealName, @RestaurantID, @RoleId, @Salt)";

        return dbConn.executeNonQuery(command) >0;
    }


 

    public bool updatePasswordByUserName()
    {
        dbConn.addParameter("@UserPassword", UserPassword);
        dbConn.addParameter("@UserName", UserName);

        string command = "UPDATE RestaurantUser SET UserPassword=@UserPassword WHERE UserName=@UserName";

        return dbConn.executeNonQuery(command) > 0; //i.e. 1 or more rows affected
    }

    public bool removeUser()
    {
        dbConn.addParameter("@UserName", UserName);
        string command = "DELETE FROM RestaurantUser WHERE UserName=@UserName";

        return dbConn.executeNonQuery(command) > 0;
    }

    public bool changeRestaurant()
    {
        bool flag;
        dbConn.addParameter("@UserName", UserName);
        dbConn.addParameter("@RestaurantID", RestaurantID);
        String command = "UPDATE RestaurantUser SET RestaurantID=@RestaurantID WHERE UserName=@UserName";

       flag = dbConn.executeNonQuery(command) > 0;

        if(flag)
        {
            HttpContext.Current.Session["RestaurantID"] = RestaurantID.ToString();
            System.Diagnostics.Debug.WriteLine("In User.ChangeRestaurant if(flag)");
        }
        System.Diagnostics.Debug.WriteLine("outside");
        return flag;

    }

    public DataTable selectAllUsers()
    {
        String command = "Select UserID,UserName FROM RestaurantUser";
        DataTable table = dbConn.executeReader(command);
        return table;
    }

    public String getUserDetails()
    {
        dbConn.addParameter("@UserName", UserName);
        String command = "Select CONCAT( RealName,  ':' ,  EmailAddress) as Details FROM RestaurantUser WHERE UserName=@UserName";
        String result="";

            DataTable table = dbConn.executeReader(command);
        if (table == null || table.Rows.Count < 1)
        {
            result = "";
        }
        else
        {

           return result = table.Rows[0]["Details"].ToString();
        }

        return result;
    }

}