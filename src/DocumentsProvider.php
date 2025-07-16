<?php
namespace Thunderkiss52\LaravelDocuments;

use Spatie\LaravelPackageTools\Package;
use Spatie\LaravelPackageTools\PackageServiceProvider;
use Spatie\LaravelPackageTools\Commands\InstallCommand;

class DocumentsProvider extends PackageServiceProvider
{
    public function configurePackage(Package $package): void
    {
        $package
            ->name('laravel-documents')
            ->hasConfigFile()
            ->hasRoute('documents')
            ->publishesServiceProvider('DocumentsProvider')
            ->hasInstallCommand(function(InstallCommand $command) {
                $command
                    ->publishConfigFile();
                    //->publishMigrations()
                    //->copyAndRegisterServiceProviderInApp();
            });
    }

    public function packageBooted(): void
    {

    }

    public function packageRegistered(): void
    {
        
    }
}
