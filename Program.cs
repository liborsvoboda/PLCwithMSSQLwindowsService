using SprayingCabineService;

var builder = Host.CreateApplicationBuilder(args);
builder.Services.AddWindowsService(cfg => { cfg.ServiceName = "Spraying Cabine Service"; }).AddHostedService<SprayingCabineService.SprayingCabineService>();

var host = builder.Build();
host.Run();
