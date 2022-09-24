<?php

namespace App\Console\Commands;

use App\Models\Sale;
use Illuminate\Console\Command;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class StoreSale extends Command
{
    /**
     * The name and signature of the console command.
     *
     * @var string
     */
    protected $signature = 'store:sales';

    /**
     * The console command description.
     *
     * @var string
     */
    protected $description = 'Enregistre les ventes et leurs acheteurs';

    /**
     * Execute the console command.
     *
     * @return int
     */
    public function handle()
    {
        $worksheet = $this->getActiveSheet(storage_path('data/SALES.xlsx'));
        $counter = 0;

        foreach ($worksheet->getRowIterator() as $row) {
            if ($counter++ === 0) continue;

            $cellIterator = $row->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(true);

            $cells = [];

            foreach($cellIterator as $cell) {
                $cells[] = $cell->getValue();
            }

            Sale::create([
                'buyer_name' => $cells[0],
                'price' => $cells[1],
            ]);

        }

        $this->comment('Ventes enregistrées avec succès');
    }

    private function getActiveSheet(string $path): Worksheet
    {
        return (new Xlsx)
            ->load($path)
            ->getActiveSheet();
    }
}
