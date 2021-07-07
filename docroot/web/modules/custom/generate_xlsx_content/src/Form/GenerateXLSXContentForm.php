<?php

namespace Drupal\generate_xlsx_content\Form;

use Drupal\Core\Entity\EntityTypeManagerInterface;
use Drupal\Core\Form\FormBase;
use Drupal\Core\Form\FormStateInterface;
use Symfony\Component\DependencyInjection\ContainerInterface;

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
   * Create an instance of GenerateXLSXContentForm.
   */
  public function __construct(EntityTypeManagerInterface $entityTypeManager) {
    $this->entityTypeManager = $entityTypeManager;
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
//    // drupal_set_message($this->t('@can_name ,Your application is being submitted!', array('@can_name' => $form_state->getValue('candidate_name'))));
//    foreach ($form_state->getValues() as $key => $value) {
//      drupal_set_message($key . ': ' . $value);
//    }
  }

}
