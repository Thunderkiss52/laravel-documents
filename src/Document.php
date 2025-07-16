<?php
namespace Thunderkiss52\LaravelDocuments;

use Illuminate\Console\Command;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Http\Request;
use Lorisleiva\Actions\ActionRequest;
use Lorisleiva\Actions\Concerns\AsAction;
use Symfony\Component\HttpFoundation\BinaryFileResponse;
abstract class Document
{
    use AsAction;

    abstract public static function label(): string;
    abstract public static function model(): string;
    abstract public function generate(Model $model): string;
    abstract public function docLabel(Model $model): string;

    protected function getTemplatePath(): string
    {
        $templateName = strtolower((new \ReflectionClass($this))->getShortName());
        return config('documents.tempate_path', storage_path('docTemplates')) . "/Document.php";
    }
    public function handle(Model $model): string
    {
        return $this->generate($model);
    }

    /*public function asJob()
    {
        
    } */

    public function getControllerMiddleware(): array
    {
        return ['auth'];
    }

    public function authorize(ActionRequest $request): bool
    {
        return true;
        //return in_array($request->user()->role, ['author', 'admin']);
    }
    public string $commandDescription = 'Сгенерировать документ';
    
    public function asCommand(Command $command)
    {
        //$user = User::findOrFail($command->argument('user_id'));
        //$this->handle($user, $command->argument('password'));
        //,$this->docLabel($model, $id);
        $mod = static::model()::findOrFail($command->argument('id'));
        $command->line("Генерация документа: {$this->docLabel($mod)}");
        $path = $this->generate($mod);
        $command->line("Файл сгенерирован и сохранён по пути: {$path}");
    }


    public function asController($id): BinaryFileResponse
    {
        return response()->download($this->generate(static::model()::findOrFail($id)),$this->docLabel(static::model()::findOrFail($id)));
    }
}