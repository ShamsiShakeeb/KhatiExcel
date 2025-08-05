using KhatiExcel.Feature;
using Microsoft.Extensions.DependencyInjection;

namespace KhatiExcel.DependencyResolver
{
    public static class ExcelFeatureResolver
    {
        public static IServiceCollection AdoDependency(this IServiceCollection services)
        {
            services.AddSingleton<ILoadExcel,LoadExcel>();
            return services;
        }
    }
}
