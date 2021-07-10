<?php

namespace Drupal\generate_xlsx_content\Form;

use Drupal\Core\Batch\BatchBuilder;
use Drupal\Core\Entity\EntityTypeManagerInterface;
use Drupal\Core\File\FileSystemInterface;
use Drupal\Core\Form\FormBase;
use Drupal\Core\Form\FormStateInterface;
use Drupal\Core\Link;
use Drupal\Core\Url;
use PhpOffice\PhpSpreadsheet\Calculation\Exception;
use Symfony\Component\DependencyInjection\ContainerInterface;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

/**
 * Implements a form to generate xlsx content.
 */
class GenerateXLSXContentForm extends FormBase {

  /**
   * The entity type manager.
   *
   * @var \Drupal\Core\Entity\EntityTypeManagerInterface
   */
  protected $entityTypeManager;

  /**
   * Batch Builder.
   *
   * @var \Drupal\Core\Batch\BatchBuilder
   */
  protected $batchBuilder;

  /**
   * Create an instance of GenerateXLSXContentForm.
   */
  public function __construct(EntityTypeManagerInterface $entityTypeManager) {
    $this->entityTypeManager = $entityTypeManager;
    $this->batchBuilder = new BatchBuilder();
  }

  /**
   * {@inheritdoc}
   */
  public static function create(ContainerInterface $container) {
    return new static(
      $container->get('entity_type.manager')
    );
  }

  /**
   * {@inheritdoc}
   */
  public function getFormId() {
    return 'generate_xlsx_content_form';
  }

  /**
   * {@inheritdoc}
   */
  public function buildForm(array $form, FormStateInterface $form_state) {
    $content_types = $this->entityTypeManager->getStorage('node_type')->loadMultiple();
    if (!empty($content_types)) {
      foreach ($content_types as $content_type) {
        $form[$content_type->id()] = [
          '#type' => 'checkbox',
          '#title' => $content_type->label(),
        ];
      }

      $form['actions']['#type'] = 'actions';
      $form['actions']['submit'] = [
        '#type' => 'submit',
        '#value' => $this->t('Generate Archive'),
        '#button_type' => 'primary',
      ];
    }

    return $form;
  }

  /**
   * {@inheritdoc}
   */
  public function validateForm(array &$form, FormStateInterface $form_state) {
    $content_types = $this->entityTypeManager->getStorage('node_type')->loadMultiple();
    foreach ($content_types as $content_type) {
      if (!empty($form_state->getValue($content_type->id()))) {
        return;
      }
    }
    $form_state->setErrorByName('content_type', $this->t('No content types was selected.'));
  }

  /**
   * {@inheritdoc}
   */
  public function submitForm(array &$form, FormStateInterface $form_state) {
    $content_types = $this->entityTypeManager->getStorage('node_type')->loadMultiple();
    if (!empty($content_types)) {
      foreach ($content_types as $content_type) {
        if (!empty($form_state->getValue($content_type->id()))) {
          $bundles[] = $content_type->id();
        }
      }

      if (!empty($bundles)) {
        $nids = \Drupal::entityQuery('node')
          ->condition('type', $bundles, 'IN')
          ->execute();

        $this->batchBuilder
          ->setTitle($this->t('Processing'))
          ->setInitMessage($this->t('Initializing.'))
          ->setProgressMessage($this->t('Completed @current of @total.'))
          ->setErrorMessage($this->t('An error has occurred.'));

        $this->batchBuilder->setFile(drupal_get_path('module', 'generate_xlsx_content') . '/src/Form/GenerateXLSXContentForm.php');
        $this->batchBuilder->addOperation([$this, 'processItems'], [$nids]);
        //$this->batchBuilder->setFinishCallback([$this, 'finished']);

        batch_set($this->batchBuilder->toArray());
      }
    }
  }

  /**
   * Processor for batch operations.
   */
  public function processItems($items, array &$context) {
    // Elements per operation.
    $limit = 50;

    // Set default progress values.
    if (empty($context['sandbox']['progress'])) {
      $context['sandbox']['progress'] = 0;
      $context['sandbox']['max'] = count($items);
    }

    // Save items to array which will be changed during processing.
    if (empty($context['sandbox']['items'])) {
      $context['sandbox']['items'] = $items;
    }

    $counter = 0;
    if (!empty($context['sandbox']['items'])) {
      // Remove already processed items.
      if ($context['sandbox']['progress'] != 0) {
        array_splice($context['sandbox']['items'], 0, $limit);
      }

      $zip_name = 'export_content_' . date('Y_m_d_H_i', \Drupal::time()->getRequestTime()) . '.zip';

      foreach ($context['sandbox']['items'] as $item) {
        if ($counter != $limit) {
          $output = $this->vboExportContentXlsx($item);
          $this->sendToFile($output, $zip_name, $item);

          $counter++;
          $context['sandbox']['progress']++;

          $context['message'] = $this->t('Now processing node :progress of :count', [
            ':progress' => $context['sandbox']['progress'],
            ':count' => $context['sandbox']['max'],
          ]);

          // Increment total processed item values. Will be used in finished
          // callback.
          $context['results']['processed'] = $context['sandbox']['progress'];
        }
      }
    }

    // If not finished all tasks, we count percentage of process. 1 = 100%.
    if ($context['sandbox']['progress'] != $context['sandbox']['max']) {
      $context['finished'] = $context['sandbox']['progress'] / $context['sandbox']['max'];
    }
  }

  /**
   * Xlsx content builder function.
   */
  protected function vboExportContentXlsx($variables) {
    //rename $variables.
    $node = $this->entityTypeManager->getStorage('node')->load($variables);
    $current_user = \Drupal::currentUser();
    $headers = [
      'Node ID',
      'Link',
      'Content Type',
      'Title',
      'Author',
      'Created at',
      'Status',
      'Uuid',
      'Author Id',
      'Langcode',
    ];

    // Load PhpSpreadsheet library.
    if (!_vbo_export_library_exists(Spreadsheet::class)) {
      \Drupal::logger('vbo_export')->error('PhpSpreadsheet library not installed.');
      return '';
    }

    // Create PHPExcel spreadsheet and add rows to it.
    $spreadsheet = new Spreadsheet();
    $spreadsheet->removeSheetByIndex(0);
    $spreadsheet->getProperties()
      ->setCreated(\Drupal::time()->getRequestTime())
      ->setCreator($current_user->getDisplayName())
      ->setTitle('VBO Export - ' . date('d-m-Y H:i', \Drupal::time()->getRequestTime()))
      ->setLastModifiedBy($current_user->getDisplayName());
    $worksheet = $spreadsheet->createSheet();
    $worksheet->setTitle((string) t('Export'));

    // Set header.
    $col_index = 1;
    foreach ($headers as $header) {
      $worksheet->setCellValueExplicitByColumnAndRow($col_index++, 1, trim($header), DataType::TYPE_STRING);
    }

    $options = ['absolute' => TRUE];
    $url = Url::fromRoute('entity.node.canonical', ['node' => $node->id()], $options);

    $rows[] = [
      'node_id' => $node->id(),
      'link' => $url->toString(),
      'content_type' => $node->bundle(),
      'title' => $node->label(),
      'author' => $node->getOwner()->name->getString(),
      'created_at' => date('d-m-Y H:i', $node->getCreatedTime()),
      'status' => $node->isPublished() ? 'published' : 'unpublished',
      'uuid' => $node->uuid(),
      'author_id' => $node->getOwnerId(),
      'langcode' => $node->get('langcode')->getString(),
    ];
    // Set rows.
    foreach ($rows as $row_index => $row) {
      $col_index = 1;
      foreach ($row as $cell) {
        $worksheet->setCellValueExplicitByColumnAndRow($col_index++, $row_index + 2, trim($cell), DataType::TYPE_STRING);
      }
    }

    // Add additional styling to the worksheet.
    $spreadsheet->getDefaultStyle()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT);
    $last_column = $worksheet->getHighestColumn();
    $last_column_index = Coordinate::columnIndexFromString($last_column);

    // Define the range of the first row.
    $first_row_range = 'A1:' . $last_column . '1';

    // Set first row in bold.
    $worksheet->getStyle($first_row_range)->getFont()->setBold(TRUE);

    // Activate an autofilter on the first row.
    $worksheet->setAutoFilter($first_row_range);

    // Set wrap text and top vertical alignment for the entire worksheet.
    $full_range = 'A1:' . $last_column . $worksheet->getHighestRow();
    $worksheet->getStyle($full_range)->getAlignment()
      ->setWrapText(TRUE)
      ->setVertical(Alignment::VERTICAL_TOP);

    for ($column = 0; $column <= $last_column_index; $column++) {
      $worksheet->getColumnDimensionByColumn($column)->setAutoSize(TRUE);
    }

    // Set a minimum and maximum width for columns.
    $min_column_width = 15;
    $max_column_width = 85;

    // Added a try-catch block
    // due to https://github.com/PHPOffice/PHPExcel/issues/556.
    try {
      $worksheet->calculateColumnWidths();
    }
    catch (Exception $e) {

    }

    for ($column = 0; $column <= $last_column_index; $column++) {
      $width = $worksheet->getColumnDimensionByColumn($column)->getWidth();
      if ($width < $min_column_width) {
        $worksheet->getColumnDimensionByColumn($column)->setAutoSize(FALSE);
        $worksheet->getColumnDimensionByColumn($column)->setWidth($min_column_width);
      }
      elseif ($width > $max_column_width) {
        $worksheet->getColumnDimensionByColumn($column)->setAutoSize(FALSE);
        $worksheet->getColumnDimensionByColumn($column)->setWidth($max_column_width);
      }
    }

    $objWriter = new Xlsx($spreadsheet);
    ob_start();
    $objWriter->save('php://output');
    $excelOutput = ob_get_clean();
    return $excelOutput;
  }

  /**
   * Output generated string to file. Message user.
   *
   * @param string $output
   *   The string that will be saved to a file.
   * @param string $zip_name
   *   File name for archive.
   * @param int $item
   *   Id of the node.
   */
  protected function sendToFile(string $output, string $zip_name, int $item) {
    if (!empty($output)) {
      $node = $this->entityTypeManager->getStorage('node')->load($item);
      $filename = 'node_' . $node->uuid() . '.xlsx';
      $wrapper = 'public';

      $destination = $wrapper . '://' . $filename;
      $file = file_save_data($output, $destination, FileSystemInterface::EXISTS_REPLACE);
      $file->setTemporary();
      $file->save();
      $file_system = \Drupal::service('file_system');
      $archiver_path = '/var/www/docroot/web/sites/default/files/' . $zip_name;

      if (!file_exists($archiver_path)) {
        $zip_file_uri = \Drupal::service('file_system')->saveData('', $archiver_path, FileSystemInterface::EXISTS_RENAME);
        $zip = \Drupal::service('plugin.manager.archiver')->getInstance(['filepath' => $file_system->realpath($zip_file_uri)])->getArchive();
        $zip->addFile($file_system->realpath($file->getFileUri()), $file->getFilename());
      }
      else {
        $zip = \Drupal::service('plugin.manager.archiver')->getInstance(['filepath' => $file_system->realpath($archiver_path)])->getArchive();
        $zip->addFile($file_system->realpath($file->getFileUri()), $file->getFilename());
      }

      $link = Link::fromTextAndUrl($this->t('Click here'), Url::fromUri('internal:/sites/default/files/' . $zip_name));
      $this->messenger()->addStatus($this->t('Export file created, @link to download.', ['@link' => $link->toString()]));
    }
  }

}
