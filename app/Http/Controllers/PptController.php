<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\IOFactory;
class PptController extends Controller
{
    public function generate()
    {
        $ppt = new PhpPresentation();
    
        // Create the first slide
        $slide1 = $ppt->getActiveSlide();
        $this->addTitleSlide($slide1, 'Slide 1 Title');
    
        // Create additional slides
        for ($i = 2; $i <= 5; $i++) {
            $slide = $ppt->createSlide();
            $this->addTitleSlide($slide, "Slide $i Title");
        }
    
        // Save the presentation
        $fileName = 'generated-presentation.pptx';
        $writer = IOFactory::createWriter($ppt, 'PowerPoint2007');
        $writer->save(storage_path('app/' . $fileName));
    
        return response()->download(storage_path('app/' . $fileName))->deleteFileAfterSend(true);
    }
    
    private function addTitleSlide(Slide $slide, $title)
    {
        $shape = $slide->createRichTextShape()
                       ->setHeight(200)
                       ->setWidth(600)
                       ->setOffsetX(100)
                       ->setOffsetY(200);
    
        $titleRun = $shape->createTextRun($title);
        $titleRun->getFont()->setBold(true);
        $titleRun->getFont()->setSize(32); // Set font size
    }
}
