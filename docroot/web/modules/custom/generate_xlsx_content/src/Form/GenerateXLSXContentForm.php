<?php

namespace Drupal\generate_xlsx_content\Form;

use Drupal\Component\Datetime\TimeInterface;
use Drupal\Core\Archiver\ArchiverManager;
use Drupal\Core\Batch\BatchBuilder;
use Drupal\Core\Datetime\DateFormatterInterface;
use Drupal\Core\Entity\EntityTypeManagerInterface;
use Drupal\Core\File\FileSystemInterface;
use Drupal\Core\Form\FormBase;
use Drupal\Core\Form\FormStateInterface;
use Drupal\Core\Link;
use Drupal\Core\Session\AccountInterface;
use Drupal\Core\Url;
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
   * The file system service.
   *
   * @var \Drupal\Core\File\FileSystemInterface
   */
  protected $fileSystem;

  /**
   * The archiver plugin manager service.
   *
   * @var \Drupal\Core\Archiver\ArchiverManager
   */
  protected $archiverManager;

  /**
   * The date formatter service.
   *
   * @var \Drupal\Core\Datetime\DateFormatterInterface
   */
  protected $dateFormatter;

  /**
   * The current user.
   *
   * @var \Drupal\Core\Session\AccountInterface
   */
  protected $currentUser;

  /**
   * The time service.
   *
   * @var \Drupal\Component\Datetime\TimeInterface
   */
  protected $time;

  /**
   * Create an instance of GenerateXLSXContentForm.
   */
  public function __construct(EntityTypeManagerInterface $entityTypeManager, FileSystemInterface $file_system, ArchiverManager $archiver_manager, DateFormatterInterface $date_formatter, AccountInterface $current_user, TimeInterface $time) {
    $this->entityTypeManager = $entityTypeManager;
    $this->batchBuilder = new BatchBuilder();
    $this->fileSystem = $file_system;
    $this->archiverManager = $archiver_manager;
    $this->dateFormatter = $date_formatter;
    $this->currentUser = $current_user;
    $this->time = $time;
  }

  /**
   * {@inheritdoc}
   */
  public static function create(ContainerInterface $container) {
    return new static(
      $container->get('entity_type.manager'),
      $container->get('file_system'),
      $container->get('plugin.manager.archiver'),
      $container->get('date.formatter'),
      $container->get('current_user'),
      $container->get('datetime.time')
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

        batch_set($this->batchBuilder->toArray());
      }
    }
  }

  /**
   * Processor for batch operations.
   */
  public function processItems($items, array &$context) {
    $limit = 50;

    if (empty($context['sandbox']['progress'])) {
      $context['sandbox']['progress'] = 0;
      $context['sandbox']['max'] = count($items);
    }

    if (empty($context['sandbox']['items'])) {
      $context['sandbox']['items'] = $items;
    }

    $counter = 0;
    if (!empty($context['sandbox']['items'])) {
      if ($context['sandbox']['progress'] != 0) {
        array_splice($context['sandbox']['items'], 0, $limit);
      }

      $zip_name = 'export_content_' . $this->dateFormatter->format($this->time->getRequestTime(), 'custom', 'Y_m_d_H_i') . '.zip';

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
          $context['results']['processed'] = $context['sandbox']['progress'];
        }
      }
    }

    if ($context['sandbox']['progress'] != $context['sandbox']['max']) {
      $context['finished'] = $context['sandbox']['progress'] / $context['sandbox']['max'];
    }
  }

  /**
   * Xlsx content builder function..
   *
   * @param int $item
   *   Id of the node.
   */
  protected function vboExportContentXlsx(int $item) {
    /** @var \Drupal\node\NodeInterface $node */
    $node = $this->entityTypeManager->getStorage('node')->load($item);
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
      $this->logger('vbo_export')->error('PhpSpreadsheet library not installed.');
      return '';
    }

    // Create PHPExcel spreadsheet and add rows to it.
    $spreadsheet = new Spreadsheet();
    $spreadsheet->removeSheetByIndex(0);
    $spreadsheet->getProperties()
      ->setCreated($this->time->getRequestTime())
      ->setCreator($this->currentUser->getDisplayName())
      ->setTitle('VBO Export - ' . $this->dateFormatter->format($this->time->getRequestTime(), 'custom', 'd-m-Y H:i'))
      ->setLastModifiedBy($this->currentUser->getDisplayName());
    $worksheet = $spreadsheet->createSheet();
    $worksheet->setTitle((string) t('Export'));

    // Set header.
    $col_index = 1;
    foreach ($headers as $header) {
      $worksheet->setCellValueExplicitByColumnAndRow($col_index++, 1, trim($header), DataType::TYPE_STRING);
    }

    $options = ['absolute' => TRUE];
    $url = Url::fromRoute('entity.node.canonical', ['node' => $node->id()], $options);

    // Array with rows.
    $rows[] = [
      'node_id' => $node->id(),
      'link' => $url->toString(),
      'content_type' => $node->bundle(),
      'title' => $node->label(),
      'author' => $node->getOwner()->name->getString(),
      'created_at' => $this->dateFormatter->format($node->getCreatedTime(), 'custom', 'd-m-Y H:i'),
      'status' => $node->isPublished() ? 'published' : 'unpublished',
      'uuid' => $node->uuid(),
      'author_id' => $node->getOwnerId(),
      'langcode' => $node->get('langcode')->getString(),
    ];

    // Set rows.
    foreach ($rows as $row_index => $row) {
      $col_index = 1;
      foreach ($row as $cell) {
        if (filter_var($cell, FILTER_VALIDATE_URL) == TRUE) {

          // Make link field clickable in xlsx table.
          $spreadsheet->getActiveSheet()->getCell('B' . $col_index)->getHyperlink()->setUrl($cell);
          // Rows start from 1 and we need to account for header.
          $worksheet->setCellValueExplicitByColumnAndRow($col_index++, $row_index + 2, trim($cell), DataType::TYPE_STRING);
        }
        else {
          // Rows start from 1 and we need to account for header.
          $worksheet->setCellValueExplicitByColumnAndRow($col_index++, $row_index + 2, trim($cell), DataType::TYPE_STRING);
        }
      }
    }

    // Add some additional styling to the worksheet.
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

    // Catch the output of the spreadsheet.
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

      /** @var \Drupal\node\NodeInterface $node */
      $node = $this->entityTypeManager->getStorage('node')->load($item);
      $filename = 'node_' . $node->uuid() . '.xlsx';
      $wrapper = 'public';

      $destination = $wrapper . '://' . $filename;
      $file = file_save_data($output, $destination, FileSystemInterface::EXISTS_REPLACE);
      $file->setTemporary();
      $file->save();
      $archiver_path = $this->fileSystem->realpath('public://' . $zip_name);

      if (!file_exists($archiver_path)) {
        $zip_file_uri = $this->fileSystem->saveData('', $archiver_path, FileSystemInterface::EXISTS_RENAME);
        $zip = $this->archiverManager->getInstance(['filepath' => $this->fileSystem->realpath($zip_file_uri)])->getArchive();
        $zip->addFile($this->fileSystem->realpath($file->getFileUri()), $file->getFilename());
      }
      else {
        $zip = $this->archiverManager->getInstance(['filepath' => $this->fileSystem->realpath($archiver_path)])->getArchive();
        $zip->addFile($this->fileSystem->realpath($file->getFileUri()), $file->getFilename());
      }
      $file_scheme = \Drupal::config('system.file')->get('default_scheme');

      // Generate link to download file with export results.
      $link = Link::fromTextAndUrl($this->t('Click here'), Url::fromUri(file_create_url($file_scheme . '://' . $zip_name)));
      $this->messenger()->addStatus($this->t('Content export file was created, @link to download.', ['@link' => $link->toString()]));
    }
  }

}
