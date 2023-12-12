namespace ExFormOfficeAddInEntities
{
    public class User
    {
        public int UserId { get; set; }
        public string FullName { get; set; }
        public char UserType { get; set; }
        public string UserName { get; set; }
        public string Email { get; set; }
        public int CompanyId { get; set; }
        public string CompanyName { get; set; }
        public bool IsActive { get; set; }
        public string Password { get; set; }
    }
}
