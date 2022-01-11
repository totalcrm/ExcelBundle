Symfony4 Excel bundle
============

### Things to know:

CSV is faster so if you have to create simple xls file,
I encourage you to use the built-in function for csv: http://php.net/manual-lookup.php?pattern=csv&lang=en&scope=quickref

## Installation

**1**  Add to composer.json to the `require` key

``` shell
    $composer require totalcrm/excel-bundle
``` 

**2** Register the bundle in ``app/AppKernel.php``

``` php
    $bundles = array(
        // ...
        new TotalCRM\ExcelBundle\TotalCRMExcelBundle(),
    );
```

## TL;DR

- Create an empty object:

``` php
$phpExcelObject = $this->get('phpexcel')->createPHPExcelObject();
```

- Create an object from a file:

``` php
$phpExcelObject = $this->get('phpexcel')->createPHPExcelObject('file.xls');
```

- Create a Excel5 and write to a file given the object:

```php
$writer = $this->get('phpexcel')->createWriter($phpExcelObject, 'Excel5');
$writer->save('file.xls');
```

- Create a Excel5 and create a StreamedResponse:

```php
$writer = $this->get('phpexcel')->createWriter($phpExcelObject, 'Excel5');
$response = $this->get('phpexcel')->createStreamedResponse($writer);
```

- Create a Excel file with an image:

```php
$writer = $this->get('phpexcel')->createPHPExcelObject();
$writer->setActiveSheetIndex(0);
$activesheet = $writer->getActiveSheet();

$drawingobject = $this->get('phpexcel')->createPHPExcelWorksheetDrawing();
$drawingobject->setName('Image name');
$drawingobject->setDescription('Image description');
$drawingobject->setPath('/path/to/image');
$drawingobject->setHeight(60);
$drawingobject->setOffsetY(20);
$drawingobject->setCoordinates('A1');
$drawingobject->setWorksheet($activesheet)
```

## Not Only 'Excel5'

The list of the types are:

1.  'Excel5'
2.  'Excel2007'
3.  'Excel2003XML'
4.  'OOCalc'
5.  'SYLK'
6.  'Gnumeric'
7.  'HTML'
8.  'CSV'

## Example

### Fake Controller

The best place to start is the fake Controller at `Tests/app/Controller/FakeController.php`, that is a working example.

### More example

You could find a lot of examples in the official PHPExcel repository https://github.com/PHPOffice/PHPExcel/tree/develop/Examples

### For lazy devs

``` php
namespace YOURNAME\YOURBUNDLE\Controller;

use Symfony\Bundle\FrameworkBundle\Controller\Controller;
use Symfony\Component\HttpFoundation\ResponseHeaderBag;

class DefaultController extends Controller
{

    public function indexAction($name)
    {
        // ask the service for a Excel5
       $phpExcelObject = $this->get('phpexcel')->createPHPExcelObject();

       $phpExcelObject->getProperties()->setCreator("totalcrm")
           ->setLastModifiedBy("Giulio De Donato")
           ->setTitle("Office 2005 XLSX Test Document")
           ->setSubject("Office 2005 XLSX Test Document")
           ->setDescription("Test document for Office 2005 XLSX, generated using PHP classes.")
           ->setKeywords("office 2005 openxml php")
           ->setCategory("Test result file");
       $phpExcelObject->setActiveSheetIndex(0)
           ->setCellValue('A1', 'Hello')
           ->setCellValue('B2', 'world!');
       $phpExcelObject->getActiveSheet()->setTitle('Simple');
       // Set active sheet index to the first sheet, so Excel opens this as the first sheet
       $phpExcelObject->setActiveSheetIndex(0);

        // create the writer
        $writer = $this->get('phpexcel')->createWriter($phpExcelObject, 'Excel5');
        // create the response
        $response = $this->get('phpexcel')->createStreamedResponse($writer);
        // adding headers
        $dispositionHeader = $response->headers->makeDisposition(
            ResponseHeaderBag::DISPOSITION_ATTACHMENT,
            'stream-file.xls'
        );
        $response->headers->set('Content-Type', 'text/vnd.ms-excel; charset=utf-8');
        $response->headers->set('Pragma', 'public');
        $response->headers->set('Cache-Control', 'maxage=1');
        $response->headers->set('Content-Disposition', $dispositionHeader);

        return $response;        
    }
}
```

## Contribute

1. fork the project
2. clone the repo
3. get the coding standard fixer: `wget http://cs.sensiolabs.org/get/php-cs-fixer.phar`
4. before the PullRequest you should run the coding standard fixer with `php php-cs-fixer.phar fix -v .`
