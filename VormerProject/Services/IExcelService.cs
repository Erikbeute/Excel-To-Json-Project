
namespace VormerProject.Services
{
    public interface IExcelService
    {
        Task<string> ReturnJson(IFormFile file);
    }
}