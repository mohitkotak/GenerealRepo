using AutoMapper;
using ClosedXML.Excel;
using GeneralWorkMVC.Models;

namespace GeneralWorkMVC.Mapper
{
    public class AutoMapperProfile : Profile
    {
        public AutoMapperProfile()
        {
            // Map User model for Sheet1
            CreateMap<IXLRow, User>()
                .ForMember(dest => dest.Name, opt => opt.MapFrom(src => src.Cell(1).Value.ToString()))
                .ForMember(dest => dest.Age, opt => opt.MapFrom(src => src.Cell(2).GetValue<int>()))
                .ForMember(dest => dest.Email, opt => opt.MapFrom(src => src.Cell(3).Value.ToString()));

            // Map Product model for Sheet2
            CreateMap<IXLRow, Product>()
                .ForMember(dest => dest.ProductName, opt => opt.MapFrom(src => src.Cell(1).Value.ToString()))
                .ForMember(dest => dest.Price, opt => opt.MapFrom(src => src.Cell(2).GetValue<decimal>()))
                .ForMember(dest => dest.Quantity, opt => opt.MapFrom(src => src.Cell(3).GetValue<int>()));
        }
    }
}