var d = require("../Database");

describe("Database", function () {
  it("should create an instance of the Database class, and set the connection string.", function () {
    var db = d.Database("my fake connection string");
    expect(typeof(db)).toBe("object");
    expect(db.ConnectionString).toBe("my fake connection string");
  });
});